Attribute VB_Name = "createOneGFReport"
Sub rty()

Dim rown As Long
Dim topFiveOneGF As String, totalOneGF As String, breadOneGF As String, chillOneGF As String, grocOneGF As String, nonRebOneGF As String
Dim wstopFiveOneGF As Worksheet ' wstotalOneGF As Worksheet, wsbreadOneGF As Worksheet, wschillOneGF As Worksheet, wsgrocOneGF As Worksheet, nonRebOneGF As Worksheet
Dim wsCSF As Worksheet
Dim copyRng As Range


Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set wsCSF = Worksheets("csfSummary")



rown = wsCSF.Cells(Rows.Count, 1).End(xlUp).Row

Set copyRng = wsCSF.Range("A1:M" & rown)

topFiveOneGF = "Top 5 Store Summary"
totalOneGF = "Overall Total 1GF Ranking"
breadOneGF = "Bread 1GF Ranking"
chillOneGF = "Chilled 1GF Ranking"
grocOneGF = "Grocery 1GF Ranking"
nonRebOneGF = "Non rebated categories Ranking"


If e(topFiveOneGF) = True Then
Worksheets(topFiveOneGF).Delete
End If

Set wstopFiveOneGF = ThisWorkbook.Sheets().Add
wstopFiveOneGF.Name = "Top 5 Store Summary"
copyRng.Copy
wstopFiveOneGF.Activate
Range("A1").PasteSpecial xlPasteValues

If e(totalOneGF) = True Then
Worksheets(totalOneGF).Delete
End If

Set wstopFiveOneGF = ThisWorkbook.Sheets().Add
wstopFiveOneGF.Name = "Overall Total 1GF Ranking"


copyRng.Copy
wstopFiveOneGF.Activate
Range("A1").PasteSpecial xlPasteValues
    
    With ActiveWorkbook.Sheets("Overall Total 1GF Ranking").Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With


modWS

If e(breadOneGF) = True Then
Worksheets(breadOneGF).Delete
End If

Set wstopFiveOneGF = ThisWorkbook.Sheets().Add
wstopFiveOneGF.Name = "Bread 1GF Ranking"
copyRng.Copy
wstopFiveOneGF.Activate
Range("A1").PasteSpecial xlPasteValues

    With ActiveWorkbook.Sheets("Bread 1GF Ranking").Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With


modWS


If e(chillOneGF) = True Then
Worksheets(chillOneGF).Delete
End If

Set wstopFiveOneGF = ThisWorkbook.Sheets().Add
wstopFiveOneGF.Name = "Chilled 1GF Ranking"
copyRng.Copy
wstopFiveOneGF.Activate
Range("A1").PasteSpecial xlPasteValues

    With ActiveWorkbook.Sheets("Chilled 1GF Ranking").Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With


modWS

If e(grocOneGF) = True Then
Worksheets(grocOneGF).Delete
End If

Set wstopFiveOneGF = ThisWorkbook.Sheets().Add
wstopFiveOneGF.Name = "Grocery 1GF Ranking"
copyRng.Copy
wstopFiveOneGF.Activate
Range("A1").PasteSpecial xlPasteValues

    With ActiveWorkbook.Sheets("Grocery 1GF Ranking").Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With

modWS


If e(nonRebOneGF) = True Then
Worksheets(nonRebOneGF).Delete
End If


Set wstopFiveOneGF = ThisWorkbook.Sheets().Add
wstopFiveOneGF.Name = "Non rebated categories Ranking"
copyRng.Copy
wstopFiveOneGF.Activate
Range("A1").PasteSpecial xlPasteValues

    With ActiveWorkbook.Sheets("Non rebated categories Ranking").Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With

modWS



End Sub

Public Function modWS()

Dim ws As Worksheet
Dim rown As Long
Dim psMas As Boolean
Dim wbName As String

wbName = ActiveWorkbook.Name

Set ws = ActiveSheet

With ws

    If ActiveSheet.Name = "Overall Total 1GF Ranking" Then
    
    .Columns("E:L").Delete
    .Columns("B:B").Delete
    
    ElseIf ActiveSheet.Name = "Bread 1GF Ranking" Then
    
    .Columns("I:M").Delete
    .Columns("F:G").Delete
    .Columns("E:E").Delete
    .Columns("B:B").Delete
        
'    psMas = Cells.Find("PSMASTERTON")
    
    ElseIf ActiveSheet.Name = "Chilled 1GF Ranking" Then
    
     .Columns("J:M").Delete
     .Columns("E:H").Delete
     .Columns("B:B").Delete
    
    
    ElseIf ActiveSheet.Name = "Grocery 1GF Ranking" Then
    
    .Columns("K:M").Delete
    .Columns("E:I").Delete
    .Columns("B:B").Delete
    
    ElseIf ActiveSheet.Name = "Non rebated categories Ranking" Then
    
    .Columns("M:M").Delete
    .Columns("E:K").Delete
    .Columns("B:B").Delete
    
    End If
    
    .Columns("D").NumberFormat = "$#,###;[Red]($#,###);0,000"
    .Range("A1:D1").AutoFilter Field:=2, Criteria1:="bake"
    rown = Cells(Rows.Count, 2).End(xlUp).Row
    .Range("$A$2:$D$" & rown).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete

    
 '   .AutoFilterMode = False
    .Range("A1").Select
 '   .Cells.Find("PSGisborne").Activate
    .Rows(ActiveCell.Row).EntireRow.Delete
    .Rows(1).Insert
    .Range("A1") = "Rank"
    .Range("B1") = "AgmtType"
    .Range("C1") = "Store"
    Dim shName As String
    shName = ActiveSheet.Name
    .Range("D1") = shName & " Total"

    lrows = Cells(Rows.Count, 2).End(xlUp).Row
    .Range("A2").Formula = "=RANK(D2,$D$2:$D$" & lrows & ",0)"
    .Range("A2:A" & lrows).FillDown
    .Calculate
    
    .Range("A1").CurrentRegion.Sort Key1:=.Range("A1:D1"), Order1:=xlAscending, _
    Header:=xlYes, OrderCustom:=1, DataOption1:=xlSortNormal
    
    
    
    .Rows(1).Insert
    wbName = Mid(wbName, 1, 3)
    shName = ActiveSheet.Name
    .Range("A1") = wbName & " " & shName
    .Range("A1:D1").Merge
    .Range("A1:D1").Interior.Color = RGB(255, 255, 0)
    .Range("A1:D1").HorizontalAlignment = xlCenter
    .Range("A1:D" & lrows + 1).Font.Name = "Calibri"
    .Range("A1:D1").BorderAround ColorIndex:=vbBlack, Weight:=xlThick
    
    With Range("A2:D" & lrows + 1).Select
        With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    End With

        
    .Columns("A:D").AutoFit
End With
End Function

