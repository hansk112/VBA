Attribute VB_Name = "createQuarterBread"
Sub createNewQuarter()


Dim custqtrWS As Worksheet






For Each custqtrWS In Sheets(Array("PSMtAlbert"))  ', "PSLincolnRd", "PSNapier", "PSTamatea", "PSBotany", "PSMtAlbert", "PSPapakura"))
custqtrWS.Activate

Dim actSheet As String
Dim rowsNumber As Long
Dim dateColnumStrt As Long
Dim agreColnum As Long


Dim colPnum As Long
Dim colQnum As Long
Dim ColVnum As Long
Dim ColWnum As Long
Dim ColXnum As Long
Dim ColYnum As Long
Dim ColAAnum As Long
Dim ColABnum As Long
Dim ColACnum As Long

Dim colP As String
Dim colQ As String
Dim ColV As String
Dim ColW As String
Dim ColX As String
Dim ColY As String
Dim ColZ As String
Dim ColAA As String
Dim ColAB As String
Dim ColAC As String


'Set customerWS = Worksheets("mapCustomer")

dateColnumStrt = ActiveSheet.Cells(11, Columns.Count).End(xlToLeft).Column
colPnum = dateColnumStrt - 10
colQnum = dateColnumStrt - 9
ColVnum = dateColnumStrt - 4
ColWnum = dateColnumStrt - 3
ColXnum = dateColnumStrt - 2
ColYnum = dateColnumStrt - 1
ColAAnum = dateColnumStrt + 1
ColABnum = dateColnumStrt + 2
ColACnum = dateColnumStrt + 3

colP = ConvertToLetter(colPnum)
colQ = ConvertToLetter(colQnum)
ColV = ConvertToLetter(ColVnum)
ColW = ConvertToLetter(ColWnum)
ColX = ConvertToLetter(ColXnum)
ColY = ConvertToLetter(ColYnum)
ColZ = ConvertToLetter(dateColnumStrt)
ColAA = ConvertToLetter(ColAAnum)
ColAB = ConvertToLetter(ColABnum)
ColAC = ConvertToLetter(ColACnum)


Range(ColZ & ":" & ColAB).EntireColumn.Insert
Range(ColZ & "10:" & ColAB & "10").Interior.ColorIndex = xlNone
Range(ColZ & "11").Select

' cut agreement dates
Range(ColV & "6:" & ColV & "8").Select

Selection.Cut
Range(ColY & "6").Select
ActiveSheet.Paste
Range(ColZ & "10").Select
Application.CutCopyMode = False

'move months
Range(ColW & "11:" & ColY & "11").Select
Selection.AutoFill Destination:=Range(ColW & "11:" & ColAB & "11"), Type:=xlFillDefault

Dim zSt As String
zSt = "Z"
Range(ColZ & "12").Formula = "=vlookup(" & ColZ & "11,dateLookup,2,false)"
Range(ColZ & "12").Select
Selection.AutoFill Destination:=Range(ColZ & "12:" & ColAB & "12"), Type:=xlFillDefault
ActiveSheet.Calculate
Range(ColZ & "12:" & ColAB & "12").Copy


Range(ColZ & "12:" & ColAB & "12").PasteSpecial xlPasteValues

'format month heading
With Range(ColAB & "10")
    .Value = "Qtr"
    .BorderAround ColorIndex:=1, Weight:=xlThin
    .Interior.Color = RGB(216, 228, 188)
End With

'sumifs formula
actSheet = ActiveSheet.Name
With Range(ColZ & "13:" & ColAB & "19," & ColZ & "23:" & ColAB & "25, " & ColZ & "28:" & ColAB & "31," & ColZ & "40:" & ColAB & "41")
    .Formula = "=SUMIFS(INDIRECT(""Table_FSNIdatabase.accdb[ExtendedPrice]"")," & _
    "INDIRECT(""Table_FSNIdatabase.accdb[StoreID]"")," & actSheet & "!$A$12," & _
    "INDIRECT(""Table_FSNIdatabase.accdb[ProductCategory]"")," & actSheet & "!$A13, " & _
    "INDIRECT(""Table_FSNIdatabase.accdb[Brand]"")," & actSheet & "!$C13," & _
    "INDIRECT(""Table_FSNIdatabase.accdb[monthText]"")," & actSheet & "!" & ColZ & "$11)"
End With

ActiveSheet.Calculate
'subtotal formula
Range(ColZ & "22").Formula = "=SUBTOTAL(9," & ColZ & "13:" & ColZ & "21)"
Range(ColZ & "22").Select
Selection.AutoFill Destination:=Range(ColZ & "22:" & ColAB & "22"), Type:=xlFillDefault

Range(ColAC & "22").Formula = "=SUM(" & colQ & "22:" & ColAB & "22)"

Range(ColZ & "27").Formula = "=SUBTOTAL(9," & ColZ & "23:" & ColZ & "26)"
Range(ColZ & "27").Select
Selection.AutoFill Destination:=Range(ColZ & "27:" & ColAB & "27"), Type:=xlFillDefault

Range(ColAC & "27").Formula = "=SUM(" & colQ & "27:" & ColAB & "27)"

Range(ColZ & "38").Formula = "=SUBTOTAL(9," & ColZ & "28:" & ColZ & "37)"
Range(ColZ & "38").Select
Selection.AutoFill Destination:=Range(ColZ & "38:" & ColAB & "38"), Type:=xlFillDefault
Range(ColAC & "38").Formula = "=SUM(" & colQ & "38:" & ColAB & "38)"

Range(ColZ & "39").Formula = "=SUBTOTAL(9," & ColZ & "13:" & ColZ & "38)"
Range(ColZ & "39").Select
Selection.AutoFill Destination:=Range(ColZ & "39:" & ColAB & "39"), Type:=xlFillDefault
Range(ColAC & "39").Formula = "=SUM(" & colQ & "39:" & ColAB & "39)"

' Rebate Percentage Calculation
Range(ColZ & "45").Formula = "=$D$45*" & ColZ & "39"
Range(ColZ & "45").Select
Selection.AutoFill Destination:=Range(ColZ & "45:" & ColAB & "45"), Type:=xlFillDefault
Range(ColAB & "47").Formula = "=" & ColZ & "45+" & ColAA & "45+" & ColAB & "45"
Range(ColAB & "48").Formula = "=" & ColAB & "47*0.15"
Range(ColAB & "49").Formula = "=" & ColAB & "47+" & ColAB & "48"
Range(ColZ & "57").Formula = "=if(" & ColZ & "39=0,0," & ColY & "61)"
Range(ColZ & "58").Formula = "=" & ColZ & "39*$D$58"
Range(ColZ & "59").Formula = "=(" & ColZ & "40*$D$59)+(" & ColZ & "41*$D$59)"
Range(ColZ & "61").Formula = "=" & ColZ & "57+" & ColZ & "58+" & ColZ & "59+" & ColZ & "60"
Range(ColZ & "57:" & ColZ & "61").Select
Selection.AutoFill Destination:=Range(ColZ & "57:" & ColAB & "61"), Type:=xlFillDefault

'csf fund section
rowsNumber = Cells(Rows.Count, 3).End(xlUp).Row
Range(ColZ & rowsNumber).Formula = "=subtotal(9," & ColZ & "65:" & ColZ & rowsNumber - 1 & ")"
Range(ColZ & rowsNumber).Select
Selection.AutoFill Destination:=Range(ColZ & rowsNumber & ":" & ColAB & rowsNumber), Type:=xlFillDefault
Range(ColZ & "60").Formula = "=" & ColZ & rowsNumber
Range(ColZ & "60").Select
Selection.AutoFill Destination:=Range(ColZ & "60:" & ColAB & "60"), Type:=xlFillDefault

'show only 12 months columns
Columns("E:" & colP).Hidden = True
Range(colQ & "13:" & colQ & "39").Borders.LineStyle = xlEdgeLeft
Columns("A").Hidden = True
Range("B1:" & ColAC & rowsNumber).Name = actSheet


Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .Orientation = xlLandscape
    .Zoom = False
    .FitToPagesTall = False
    .FitToPagesWide = False
    .DifferentFirstPageHeaderFooter = False
    .RightHeaderPicture.Filename = "M:\GF New Zealand\Finance\Grocery Sales\Grocery Reporting\Hans Adhoc\Rebate Project\gfLogo.gif"
    .RightHeader = "&G"
    .PrintTitleRows = "$1:$1"
End With
Application.PrintCommunication = True

ActiveSheet.Rows(62).PageBreak = xlPageBreakManual

Next custqtrWS


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


















