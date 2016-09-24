Attribute VB_Name = "getMerch"
Sub Unhide_Columns()
    'Excel objects.
    
Dim wb As Workbook
Dim wsc As Worksheet

'Set wb = Workbooks("control").Worksheets("control")

Application.ScreenUpdating = False
Application.DisplayAlerts = False

getPath

  
If e("data") = True Then
Worksheets("data").Delete
End If

Sheets.Add.Name = "data"
    With ActiveSheet
        .Range("A1") = "Region"
        .Range("B1") = "SheetName"
        .Range("C1") = "PathName"
        .Range("D1") = "FileName"
        .Range("E1") = "FMS"
        .Range("F1") = "Run"
        .Range("G1") = "Store"
        .Range("H1") = "GF Employee"
        .Range("I1") = "Contractor"
        .Range("J1") = "Shift"
        .Range("K1") = "Bread"
        .Range("L1") = "Milk"
        .Range("M1") = "Chilled/Pies/Frozen"
        .Range("N1") = "Home Ingredients/Ernesst Adams"
        .Range("O1") = "Total"
        .Range("P1") = "Unpaid Breaks"
        .Range("Q1") = "Paid Breaks"
        .Range("R1") = "Travel"
        .Range("S1") = "Kms"
        .Range("T1") = "Dairy"

    End With

Worksheets("control").Activate
rowpath = Cells(Rows.Count, 2).End(xlUp).row
Dim arrFiles() As Variant
Dim arrFileName() As Variant
arrFiles() = Range("B2:B" & rowpath)
arrFileName() = Range("A2:A" & rowpath)
Debug.Print arrFiles(3, 1)
rowpath = rowpath - 1
' workbook path open
For p = 1 To rowpath
    Workbooks.Open (arrFiles(p, 1))
    
    ' workbook file activate
    Workbooks(arrFileName(p, 1)).Activate

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
   ' ws.Name

'get last row
    Dim rownum As Long
    Dim shift As String
    Dim m_wbBook As Workbook
    Dim m_wsSheet As Worksheet
    Dim m_rnCheck As Range
    Dim m_rnFind As Range
    Dim m_stAddress As String

    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
   
    'Initialize the Excel objects.
    Set m_wbBook = ThisWorkbook
    Set m_wsSheet = m_wbBook.ActiveSheet
    
    'Search the four columns for any constants.
    Set m_rnCheck = m_wsSheet.Columns("A:AA").SpecialCells(xlCellTypeConstants)
    Dim mName As String
    With m_rnCheck
    Set m_rnFind = .find(what:="Merchandisers")
        If Not m_rnFind Is Nothing Then
            mName = CStr(m_rnFind.Address)
            Debug.Print m_rnFind.Address
            'Unhide the column, and then find the next X.
            Do
         '       m_rnFind.EntireColumns ("A:D")
                i = i + 1
                Set m_rnFind = .FindNext(m_rnFind)
                
                rng = CStr(m_rnFind.Address)
                mName = Range(rng).Value
                Debug.Print m_rnFind.Address & nRows & " " & ncols & rng
            Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress
        End If
    End With

        Workbooks(arrFileName(p, 1)).Activate
    
    shift = Range("B1")
    mrun = Range("C1")
    Cells.find("Stores").Offset(1, 0).Activate
    ActiveCell.End(xlDown).Select
    rownum = ActiveCell.row


    For k = 3 To rownum
        shName = ActiveSheet.Name
        storename = Range("B" & k)
        bread = Workbooks(arrFileName(p, 1)).ActiveSheet.Range("B" & k).Offset(0, 1)
        ernst = Workbooks(arrFileName(p, 1)).ActiveSheet.Range("B" & k).Offset(0, 2)
        chill = Workbooks(arrFileName(p, 1)).ActiveSheet.Range("B" & k).Offset(0, 3)
        dairy = Workbooks(arrFileName(p, 1)).ActiveSheet.Range("B" & k).Offset(0, 4)
        break = Workbooks(arrFileName(p, 1)).ActiveSheet.Range("B" & k).Offset(0, 5)
        travel = Workbooks(arrFileName(p, 1)).ActiveSheet.Range("B" & k).Offset(0, 6)
        km = Workbooks(arrFileName(p, 1)).ActiveSheet.Range("B" & k).Offset(0, 8)
        

        lastRowCont = Workbooks("control").Worksheets("data").Cells(Rows.Count, 6).End(xlUp).row
        lastRowCont = lastRowCont + 1
        Workbooks("control").Worksheets("data").Range("B" & lastRowCont) = shName
        Workbooks("control").Worksheets("data").Range("C" & lastRowCont) = arrFiles(p, 1)
        Workbooks("control").Worksheets("data").Range("D" & lastRowCont) = arrFileName(p, 1)
        Workbooks("control").Worksheets("data").Range("F" & lastRowCont) = mrun
        Workbooks("control").Worksheets("data").Range("G" & lastRowCont) = storename
        Workbooks("control").Worksheets("data").Range("H" & lastRowCont) = mName
        Workbooks("control").Worksheets("data").Range("J" & lastRowCont) = shift
        Workbooks("control").Worksheets("data").Range("K" & lastRowCont) = bread
        Workbooks("control").Worksheets("data").Range("N" & lastRowCont) = ernst
        Workbooks("control").Worksheets("data").Range("M" & lastRowCont) = chill
        Workbooks("control").Worksheets("data").Range("T" & lastRowCont) = dairy
        Workbooks("control").Worksheets("data").Range("Q" & lastRowCont) = break
        Workbooks("control").Worksheets("data").Range("R" & lastRowCont) = travel
        Workbooks("control").Worksheets("data").Range("S" & lastRowCont) = km
        
        
        
        
    Next k


Next ws
ActiveWorkbook.Close savechanges:=False


Next p

Workbooks("control").Worksheets("data").Activate


Columns("A:T").AutoFit

End Sub


Sub getMerch()

getPath

Dim ws As Worksheet
Dim wb As Workbook
Dim rng As Range
Dim rngFind As Range
Dim nameFind As String

Set wb = ThisWorkbook
Set ws = wb.ActiveSheet

'search range
Set rng = ws.Range("A1:D20").SpecialCells(xlCellTypeConstants)

'retrieve merch from range
With rng
    Set rngFind = .find(what:="X")
    If Not rngFind Is Nothing Then
        nameFind = rngFind.Address
        Debug.Print rngFind.Address
        Do

        Set rngFind = .FindNext(rngFind)
        Debug.Print rngFind.Address
        Loop While rngFind Is Nothing And rngFind.Address <> nameFind
    
    End If


End With


End Sub


Sub sheetray()
Dim SNarray, i
ReDim SNarray(1 To Sheets.Count)
For i = 1 To Sheets.Count
SNarray(i) = ThisWorkbook.Sheets(i).Name
Debug.Print SNarray(i)
Next
End Sub

Function getPath()
' find all files in folder
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim subfolder As Object
Dim i As Integer
Dim actPath As String
Dim wsSalesForce As Worksheet

Application.ScreenUpdating = False
Application.DisplayAlerts = False

If e("control") = True Then
Worksheets("control").Delete
End If
Sheets.Add.Name = "control"

Set objFSO = CreateObject("Scripting.FileSystemObject")

'actPath = Application.ActiveWorkbook.Path
actPathDunedin = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Aitken John\Dunedin"
actPathGore = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Aitken John\Gore"
actpathOamaru = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Aitken John\Oamaru"
actAshburton = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Angela Durham\Ashburton"
actChristchurch = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Angela Durham\Christchurch"
actTimaru = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Angela Durham\Timaru"
actCentOtago = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Cullen\Central Otago & Queenstown"
actInvercagill = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Cullen\Invercargill & Winton"
actChch = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Edkins\Christchurch"
actWestCoast = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Edkins\Westcoast"
actBlenheim = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Blenheim"
actKaikoura = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Kaikoura"
actMotueka = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Motueka"
actNelson = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Nelson"
actWestport = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Westport"

Dim pathArray(0 To 14) As String
pathArray(0) = actPathDunedin
pathArray(1) = actPathGore
pathArray(2) = actpathOamaru
pathArray(3) = actAshburton
pathArray(4) = actChristchurch
pathArray(5) = actTimaru
pathArray(6) = actCentOtago
pathArray(7) = actInvercagill
pathArray(8) = actChch
pathArray(9) = actWestCoast
pathArray(10) = actBlenheim
pathArray(11) = actKaikoura
pathArray(12) = actMotueka
pathArray(13) = actNelson
pathArray(14) = actWestport



Set objFolder = objFSO.getfolder(actPathDunedin)
i = 1

'"""""""""""""""""""""""""""""""""""""""""""""'
'""""' name range of ACTIVE SHEETS and name range of customers for below loop """""
'"""""""""""""""""""""""""""""""""""""""""""""'

Cells(1, 1).Value = "fileName"
Cells(1, 2).Value = "filePath"
Cells(1, 3).Value = "folderName"
Cells(1, 4).Value = "region"

For p = 0 To 14
Set objFolder = objFSO.getfolder(pathArray(p))

For Each objFile In objFolder.Files
Dim rgn As String
Cells(i + 1, 1) = objFile.Name
Cells(i + 1, 2) = objFile.Path
Cells(i + 1, 3) = pathArray(p)
'rgn = Find(pathArray(p))

i = i + 1
Next objFile
p = p + 1
Next p

'i = 1
'
'For Each subfolder In objFolder.subfolders
'Cells(i + 1, 3) = subfolder.Name
'Cells(i + 1, 4) = subfolder.Path
'
'i = i + 1
'Next subfolder
'ps = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Aitken John"
'
'i = 1
'For Each subfolder In objFolder.subfolders
'
'Cells(i + 1, 5) = subfolder.Name
'Cells(i + 1, 6) = subfolder.Path
'i = i + 1
'Next subfolder

ActiveSheet.Columns("A:D").AutoFit


End Function

Sub test()

Dim rownum As Long
Cells.find("Stores").Offset(1, 0).Activate
ActiveCell.End(xlDown).Select
rownum = ActiveCell.row
Debug.Print row

For k = 4 To rownum
    storename = Range("B" & k)
    
'    If storename = "travel" Or storename = "TRAVEL" Then
'    Debug.Print "Y"
'    ElseIf storename = "break" Then
'
'
'    End If
    
    bread = Range("B" & k).Offset(0, 1)
    ernst = Range("B" & k).Offset(0, 2)
    chill = Range("B" & k).Offset(0, 3)
    dairy = Range("B" & k).Offset(0, 4)
    break = Range("B" & k).Offset(0, 5)
    travel = Range("B" & k).Offset(0, 6)
    
    Debug.Print storename & " " & travel
    
Next k

End Sub


Sub testU()

'actPath = Application.ActiveWorkbook.Path
actPathDunedin = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Aitken John\Dunedin"
actPathGore = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Aitken John\Gore"
actpathOamaru = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Aitken John\Oamaru"
actAshburton = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Angela Durham\Ashburton"
actChristchurch = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Angela Durham\Christchurch"
actTimaru = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Angela Durham\Timaru"
actCentOtago = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Cullen\Central Otago & Queenstown"
actInvercagill = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Cullen\Invercargill & Winton"
actChch = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Edkins\Christchurch"
actWestCoast = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Chris Edkins\Westcoast"
actBlenheim = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Blenheim"
actKaikoura = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Kaikoura"
actMotueka = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Motueka"
actNelson = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Nelson"
actWestport = "C:\Users\Hans.Kalders\Documents\South Island Merch Run\Gaya Lowe\Westport"

Dim pathArray(0 To 14) As String
pathArray(0) = actPathDunedin
pathArray(1) = actPathGore
pathArray(2) = actpathOamaru
pathArray(3) = actAshburton
pathArray(4) = actChristchurch
pathArray(5) = actTimaru
pathArray(6) = actCentOtago
pathArray(7) = actInvercagill
pathArray(8) = actChch
pathArray(9) = actWestCoast
pathArray(10) = actBlenheim
pathArray(11) = actKaikoura
pathArray(12) = actMotueka
pathArray(13) = actNelson
pathArray(14) = actWestport




Worksheets("control").Activate
rowpath = Cells(Rows.Count, 2).End(xlUp).row
Dim arrFiles() As Variant
arrFiles() = Range("B2:B" & rowpath)
Debug.Print arrFiles(3, 1)
Workbooks.Open (arrFiles(6, 1))
MsgBox ("do something")

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
MsgBox (ws.Name)

Next ws



ActiveWorkbook.Close savechanges:=False
Worksheets("control").Activate

For i = 2 To rowpath
    


Next


End Sub
