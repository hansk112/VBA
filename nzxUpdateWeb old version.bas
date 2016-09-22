Attribute VB_Name = "Module2"
 Sub incomeStatement(scode As String)

 
   With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://nz.finance.yahoo.com/q/is?s=" & scode & ".NZ&annual", Destination:=Range( _
        "$A$1"))
        .Name = "is?s=" & scode & ".NZ&annual"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "7,8,9,10"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With

End Sub
Sub balancesheet(scode As String)

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://nz.finance.yahoo.com/q/bs?s=" & scode & ".NZ&annual", Destination:=Range( _
        "$H$1"))
        .Name = "bs?s=" & scode & ".NZ&annual"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "6,7,8,9"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub cashFlow(scode As String)

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://nz.finance.yahoo.com/q/cf?s=" & scode & ".NZ&annual", Destination:=Range( _
        "$O$1"))
        .Name = "cf?s=" & scode & ".NZ&annual"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "7,8,9,10"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub finRatios(scode As String)


    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://nz.finance.yahoo.com/q/ks?s=" & scode & ".NZ", Destination:=Range("$V$1"))

        .Name = "ks?s=" & scode & ".NZ"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "7,8,10,11,13,15,17,19,21,23"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub recoStock(scode As String)

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://nz.finance.yahoo.com/q/ao?s=" & scode & ".NZ", Destination:=Range("$AC$1"))

        .Name = "ao?s=" & scode & ".NZ"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "7,8,9,10,11,13,14,15,17"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub keyStock(scode As String)

    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://nz.finance.yahoo.com/q?s=" & scode & ".NZ", Destination:=Range("$AH$1"))

        .Name = "q?s=" & scode & ".NZ"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = """table1"",""table2"",5,6"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub formatSheet(scode As String, cName As String, sName As String)


With ActiveSheet
    .Range("A42:A46").ClearContents
    .Range("A1") = cName
    .Range("A2") = scode
    .Range("B2") = sName
    .Range("O34:O38").ClearContents
    .Range("AC1:AC5").ClearContents
    .Range("C6").Formula = "=C5/C4"
    .Range("C6:F6").FillRight
    
    
End With

End Sub


Sub recommendSum()
Dim sRec As String
Dim wsRec As Worksheet, ws1 As Worksheet

sRec = "Summary Recommend"

   If checkExists(sRec) = True Then
   Worksheets(sRec).Delete
   End If
   
   ThisWorkbook.Sheets.Add.Name = sRec
   
   Set wsRec = Worksheets("Summary Recommend")
   Set ws1 = Worksheets("list")
   
   wsRec.Range("A1") = "Stock"
   wsRec.Range("B1") = "Cur Mth"
   wsRec.Range("C1") = "Last Mth"
   wsRec.Range("D1") = "Two Mth"
   wsRec.Range("E1") = "Three Mth"
   
   
   nrowsList = ws1.Cells(Rows.Count, 1).End(xlUp).Row
   
   For i = 1 To nrowsList
    ws1.Activate
    
    code = ws1.Range("A" & i)
    Worksheets(code).Activate
     
    With ActiveSheet
        .Range("AC20") = code
        .Range("AC20:AG24").Copy
    End With
    wsRec.Activate
    
    nrows = wsRec.Cells(Rows.Count, 1).End(xlUp).Row
    wsRec.Range("A" & nrows + 1).PasteSpecial xlPasteValues
   
   
   Next i
   
   wsRec.Activate
    nrows = wsRec.Cells(Rows.Count, 1).End(xlUp).Row
    Range("F2").Formula = "=IFERROR(VLOOKUP(A2,list!A:B,2,FALSE),"""")"
   wsRec.Range("F2:F" & nrows).FillDown
   wsRec.Range("F2:F" & nrows).Copy
   wsRec.Range("F2:F" & nrows).PasteSpecial xlPasteValues
   Application.CutCopyMode = False
   wsRec.Range("G1") = "Company Name"


   wsRec.Range("g2").Formula = "=IF(F2="""",g1,F2)"
   wsRec.Range("g2:g" & nrows).FillDown
   wsRec.Range("g2:g" & nrows).Copy
   wsRec.Range("g2:g" & nrows).PasteSpecial xlPasteValues
   wsRec.Columns("F:F").Delete
  wsRec.Range("A1:F1").AutoFilter
   wsRec.Columns("a:g").AutoFit
   
   wsRec.Range("A1").Select
   
End Sub
