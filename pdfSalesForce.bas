Attribute VB_Name = "pdfSalesForce"
Sub pdfSalesForce()

””””””””””””””””””””””””””
' New format for Sales Force with Upper Case”””
””””””””””””””””””””””””””
Dim nRows As Long
Dim vPage As Long
Dim nCols As Long
Dim actName As String
Dim path As String
Dim ws As Worksheet

ActiveSheet.Range(“A1”).Select

path = ActiveWorkbook.path
path = path

For Each ws In Worksheets
ws.Activate
ws.Range(“A1”).Select

If ws.Name = “mapCustomer” Or ws.Name = “salesforce” Or ws.Name = “csfInvoices” Then GoTo unPDFSheets:

ws.PageSetup.Orientation = xlLandscape
actName = ActiveSheet.Name

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=actName, _
Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
:=False, OpenAfterPublish:=False

unPDFSheets:

Next ws

Call di

End Sub

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

Sub di()
' find all files in folder
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer
Dim actPath As String
Dim wsSalesForce As Worksheet

Sheets.Add.Name = “salesforce”

Set objFSO = CreateObject(“Scripting.FileSystemObject”)
actPath = Application.ActiveWorkbook.path
Set objFolder = objFSO.getfolder(actPath)
i = 1

‘wsSalesForce = Worksheets(“SalesForceUpload”)

”””””””””””””””””””””””””””””””””””””””””””””’
' name range of ACTIVE SHEETS and name range of customers for below loop ”””””
”””””””””””””””””””””””””””””””””””””””””””””’

Cells(1, 1).Value = “fileName”
Cells(1, 2).Value = “filePath”
Cells(1, 3).Value = “sheetName”
Cells(1, 4).Value = “salesForceID”
Cells(1, 5).Value = “customerHierarchy3”

For Each objFile In objFolder.Files
' print file name
Cells(i + 1, 1) = objFile.Name
' print file path
Cells(i + 1, 2) = objFile.path
'  trim path to sheet name
Cells(i + 1, 3).Formula = "=TRIM(MID(A” & j + 2 & “,1,LEN(A” & j + 2 & “)-4))”"
' look up sales force id
Cells(i + 1, 4).Formula = "=INDEX(mapCustomer!$AX$4:$AX$43,MATCH(salesForce!C” & j + 2 & “,mapCustomer!$A$4:$A$43,0))”"
' look up ch3 id
Cells(i + 1, 5).Formula = "=VLOOKUP(C” & j + 2 & “,mapCustomer!$A$2:$D$43,4,FALSE)”"
i = i + 1
j = j + 1
Next objFile

Range(“A1”).Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.AutoFilter
nRowsSF = Cells(Rows.Count, 4).End(xlUp).Row

ActiveSheet.Range("$A$1:$E$" & nRowsSF).AutoFilter Field:=4, Criteria1:="#N/A"
Range(“D2”).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.EntireRow.Delete
ActiveSheet.AutoFilterMode = False

' copy paste region
Range(“A1”).CurrentRegion.Copy
Range(“A1”).PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range(“A1”).Select

Dim wb As Workbook
Set wb = Workbooks.Add

ThisWorkbook.Sheets(“salesforce”).Copy before:=wb.Sheets(1)
wb.SaveAs “salesforce”

End Sub

Sub rt()

Range(“A1”).CurrentRegion.Copy
Range(“A1”).PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range(“A1”).Select

End Sub



