Attribute VB_Name = "Module1"
Sub updateHierarchy()

Dim hierachy As String

Dim wb As Workbook
Dim cWB As ThisWorkbook, custWB As Workbook
Dim nrowsR As Long, nrowsG As Long, nrowsCH As Long

Set cWB = ThisWorkbook

Application.ScreenUpdating = False
Application.DisplayAlerts = False

For Each ws In ThisWorkbook.Worksheets
    If ws.Name = "CustomerHierarchy" Then
        ws.Delete
    End If

Next ws


hierachy = "M:\GF New Zealand\Finance\Grocery Sales\Grocery Reporting\Hans Adhoc\Customer Hierarchy\Customer Hierarchy.xlsm"


'Set wb = Workbooks("Customer Hierarchy.xlsm")
Workbooks.Open (hierachy)
Set custWB = Workbooks("Customer Hierarchy.xlsm")
'wb.Activate
Worksheets("RFS").Activate
nrowsR = Cells(Rows.Count, 8).End(xlUp).Row
With ActiveSheet
    .Range("G15:S" & nrowsR).Copy

End With

cWB.Activate
'ThisWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
Set wstopFiveOneGF = ThisWorkbook.Sheets().Add
wstopFiveOneGF.Name = "CustomerHierarchy"

cWB.Worksheets("CustomerHierarchy").Activate
Worksheets("CustomerHierarchy").Range("A1").PasteSpecial xlPasteValues
custWB.Worksheets("Grocery").Activate
nrowsG = Cells(Rows.Count, "G").End(xlUp).Row

With ActiveSheet
    .Range("G16:S" & nrowsG).Copy
End With

cWB.Worksheets("CustomerHierarchy").Activate
nrowsCH = Worksheets("CustomerHierarchy").Cells(Rows.Count, 8).End(xlUp).Row
cWB.Worksheets("CustomerHierarchy").Range("A" & nrowsG + 1).PasteSpecial xlPasteValues
cWB.Worksheets("CustomerHierarchy").Columns("A:M").AutoFit

custWB.Close



Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

