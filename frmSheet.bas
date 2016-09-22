Attribute VB_Name = "frmSheet"
Sub pgSetup()

Dim nrows As Long

nrows = Cells(Rows.Count, 26).End(xlUp).Row
With ActiveSheet.PageSetup
    .PrintArea = "$Q$45:$AS$" & nrows
    ActiveSheet.Columns("AU:BA").AutoFit
    ActiveSheet.Columns("S:AS").Font.Size = "14"
    .Orientation = xlLandscape
    .PaperSize = xlPaperA3
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    

End With
ActiveSheet.Columns("S:As").AutoFit
ActiveSheet.Columns("U").ColumnWidth = 15

End Sub

Sub copyPaste()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Cells.Copy
ws.Cells.PasteSpecial xlPasteValues
'Application.CutCopyMode = False
'Range("T44").Select

Next ws

End Sub


Sub hidBelowSalary()

Dim mthSal As String
Dim delRow As Long

mthSal = "Monthly Salary"

nrows = Cells(Rows.Count, 17).End(xlUp).Row

Cells.Find(mthSal).Activate
delRow = ActiveCell.Row

Range(delRow & ":" & nrows).EntireRow.Delete


End Sub

Sub pgSetupWages()



Dim nrows As Long

nrows = Cells(Rows.Count, 26).End(xlUp).Row
With ActiveSheet.PageSetup
    .PrintArea = "$P$45:$AC$" & nrows
'    ActiveSheet.Columns("R:AC").AutoFit
'    ActiveSheet.Columns("R:AC").Font.Size = "14"
    .Orientation = xlLandscape
    .PaperSize = xlPaperA3
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    

End With
'ActiveSheet.Columns("P:AC").AutoFit
'ActiveSheet.Columns("U").ColumnWidth = 15
End Sub
