Attribute VB_Name = "crtCSFSummary"
Sub createCSF()



Dim mth As String
Dim csfBake As Long
Dim csfChill As Long
Dim arrStoreID() As Variant
Dim agmtType() As Variant
Dim wsName() As Variant
Dim arrPayFreq As Variant
Dim i As Integer
Dim rw As Range
Dim lastrow As Long
Dim ws As Worksheet
Dim wsCSFName As Worksheet
Dim numSt As String

arrStoreID = Range("sapID")
agmtType = Range("agmtType")
wsName = Range("wsName")
arrPayFreq = Range("payFreq")

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'mth = InputBox("Enter month")


mth = "21/02/2016 - 26/03/2016"

Debug.Print arrStoreID(1, 1)

numSt = "csfSummary"

If e(numSt) = True Then
Worksheets(numSt).Delete
End If

Set wsCSFName = ThisWorkbook.Sheets().Add
wsCSFName.Name = "csfSummary"

Range("A1") = "StoreID"
Range("B1") = "AgmtType"
Range("C1") = "wsName"
Range("D1") = "payFreq"
Range("E1") = "BakeCSF"
Range("F1") = "ChillCSF"
Range("G1") = "BakeTotal"
Range("H1") = "ChillTotal"
Range("I1") = "GrocTotal"
Range("J1") = "RebTotal"
Range("K1") = "OthTotal"
Range("L1") = "GrndTotal"
Range("M1") = "BakReb"
Range("N1") = "BkLowRel"
Range("O1") = "ChilReb"
Range("P1") = "YogReb"
Range("Q1") = "CulReb"
Range("R1") = "GroReb"
Range("S1") = "OneGFReb"
Range("T1") = "Total"
Range("U1") = "homBrd"
Range("V1") = "loafBrd"
Range("W1") = "occBBrd"
Range("X1") = "cultBrd"
Range("Y1") = "everBrd"
Range("Z1") = "fresBBrd"
Range("AA1") = "specCBrd"
Range("AB1") = "spreBrd"
Range("AC1") = "yogDBrd"
Range("AD1") = "flouMBrd"
Range("AE1") = "oilDBrd"
Range("AF1") = "sweBBrd"
Range("AG1") = "uhtBBrd"
Range("AH1") = "stoBBrd"
Range("AI1") = "frozPBrd"
Range("AJ1") = "conMBrd"
Range("AK1") = "valLBrd"
Range("AL1") = "buttBrd"
Range("AM1") = "fwmBrd"
Range("AN1") = "Brnd"
Range("AT1") = "Chk"
Range("AO1") = "LoafChk"
Range("AP1") = "ChilChk"
Range("AQ1") = "GrocChk"
Range("AR1") = "OthChk"
Range("AS1") = "GrndChk"


'loop storeID
For i = 1 To UBound(arrStoreID)
   Range("A" & i + 1) = arrStoreID(i, 1)
Next i

'loop agmtTyoe
For i = 1 To UBound(agmtType)
    Range("B" & i + 1) = agmtType(i, 1)
Next i

For i = 1 To UBound(wsName)
    Range("C" & i + 1) = wsName(i, 1)
Next i

For i = 1 To UBound(arrPayFreq)
    Range("D" & i + 1) = arrPayFreq(i, 1)
Next i

' delete blank row
Range("A1").Select
Selection.SpecialCells(xlLastCell).Select
lastrow = ActiveCell.Row
Range("A1:A" & lastrow).Select


With Application
    .ScreenUpdating = False
    Selection.SpecialCells(xlCellTypeBlanks).Select
    For Each rw In Selection.Rows
        If WorksheetFunction.CountA(Selection.EntireRow) = 0 Then
            Selection.EntireRow.Delete
        End If
    
    Next rw

End With

' Total Branded Loaf & Occ Bake
'Range("B2").Formula = "="
bakeBrd = "Total Branded Loaf & Occ Bake"
bake = "Baking - Category Support Fund"
chill = "Chilled - Category Support Fund"
closBal = "Closing Balance"
bakeTotal = "Baking Total"
chillTotal = "Chilled Total"
GrocTotal = "Grocery Total"
RebateTotal = "Rebate Total"
othTotal = "Total Other Rebate"
grndTotal = "Grand Total "

'brand ROW AND BRAND STRING 1gf
homBrd = "Home Bakery"
loafBrd = "Loaf"
occBBrd = "Occasion Bakery"
cultBrd = "Cultured Foods"
everBrd = "Everyday Cheese"
fresBBrd = "Fresh Beverages"
specCBrd = "Speciality Cheese"
spreBrd = "Spreads"
yogDBrd = "Yoghurt & Dairy Food"
flouMBrd = "Flour & Mixes"
oilDBrd = "Oils, Dressing & May"
sweBBrd = "Sweet Bake"
uhtBBrd = "UHT Beverages"
stoBBrd = "Store Bake"
frozPBrd = "Frozen Pastry"
conMBrd = "Convenience Meals"
valLBrd = "Value Loaf"
buttBrd = "Butters"
fwmBrd = "Fresh White Milk"

'baking
loaBrdBake = "Branded Loaf NIV"
homBrdBake = "Branded Home Bakery NIV"
occBrdBake = "Branded Occasion Bakery NIV"

' get storeid from array

Dim k As Integer
k = 2
i = 1

For k = 2 To lastrow
    'arr num
   
    
    Debug.Print k
    Debug.Print i
    
    Dim wsCSFSh As String
    Dim cCSFRow As Long
    Dim bCSFRow As Long
    Dim mthCOl As Long
    Dim valCSFBk As String
    Dim valCSFCh As String
    Dim totBakBRow As Long
    Dim totBakRow As Long
    Dim totChlRow As Long
    Dim totGroRow As Long
    Dim othRebRow As Long
    Dim totRebRow As Long
    Dim totGrdRow As Long
    Dim totBakB As String
    Dim totBak As String
    Dim totChl As String
    Dim totGro As String
    Dim totOth As String
    Dim totReb As String
    Dim totGrndReb As String
    'brand object 1gf
    Dim valhomBrd As String
    Dim valloafBrd As String
    Dim valoccBBrd As String
    Dim valcultBrd As String
    Dim valeverBrd As String
    Dim valfresBBrd As String
    Dim valspecCBrd As String
    Dim valspreBrd As String
    Dim valyogDBrd As String
    Dim valflouMBrd As String
    Dim valoilDBrd As String
    Dim valsweBBrd As String
    Dim valuhtBBrd As String
    Dim valstoBBrd As String
    Dim valfrozPBrd As String
    Dim valconMBrd As String
    Dim valvalLBrd As String
    Dim valbuttBrd As String
    Dim valfwmBrd As String
    Dim rowhomBrd As Long
    Dim rowloafBrd As Long
    Dim rowoccBBrd As Long
    Dim rowcultBrd As Long
    Dim roweverBrd As Long
    Dim rowfresBBrd As Long
    Dim rowspecCBrd As Long
    Dim rowspreBrd As Long
    Dim rowyogDBrd As Long
    Dim rowflouMBrd As Long
    Dim rowoilDBrd As Long
    Dim rowsweBBrd As Long
    Dim rowuhtBBrd As Long
    Dim rowstoBBrd As Long
    Dim rowfrozPBrd As Long
    Dim rowconMBrd As Long
    Dim rowvalLBrd As Long
    Dim rowbuttBrd As Long
    Dim rowfwmBrd As Long
    
    'brand object bakin
    Dim valloaBrdBake As String
    Dim valhomBrdBake As String
    Dim valoccBrdBake As String
    Dim rowloaBrdBake As Long
    Dim rowhomBrdBake As Long
    Dim rowoccBrdBake As Long
        
    ' need to reference
     wsCSFSh = Worksheets("csfSummary").Range("C" & k)
     Worksheets(wsCSFSh).Activate
    'find storeid
    With ActiveSheet
        If agmtType(i, 1) = "chilled" And arrPayFreq(i, 1) = "Qtr" Then
        
        Cells.Find(What:=chill).Activate
        ActiveCell.End(xlUp).Select
        cCSFRow = ActiveCell.Row
        
        Cells.Find(What:=mth).Activate
        mthCOl = ActiveCell.Column
        
        Cells.Find(What:=chillTotal).Activate
        totChlRow = ActiveCell.Row
        
        Cells.Find(What:=GrocTotal).Activate
        totGroRow = ActiveCell.Row
        
        Cells.Find(What:=othTotal).Activate
        othRebRow = ActiveCell.Row
        
        Cells.Find(What:=RebateTotal).Activate
        totRebRow = ActiveCell.Row
        
        Cells.Find(What:=grndTotal).Activate
        totGrdRow = ActiveCell.Row
        
        
        Cells.Find(What:=cultBrd).Activate
        rowcultBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=everBrd).Activate
        roweverBrd = ActiveCell.Offset(2, 0).Row
        
        Cells.Find(What:=fresBBrd).Activate
        rowfresBBrd = ActiveCell.Offset(10, 0).Row

        Cells.Find(What:=specCBrd).Activate
        rowspecCBrd = ActiveCell.Offset(3, 0).Row

        Cells.Find(What:=spreBrd).Activate
        rowspreBrd = ActiveCell.Offset(5, 0).Row
        
        Cells.Find(What:=yogDBrd).Activate
        rowyogDBrd = ActiveCell.Offset(4, 0).Row

        Cells.Find(What:=flouMBrd).Activate
        rowflouMBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=oilDBrd).Activate
        rowoilDBrd = ActiveCell.Offset(5, 0).Row
        
        Cells.Find(What:=sweBBrd).Activate
        rowsweBBrd = ActiveCell.Offset(4, 0).Row

        Cells.Find(What:=uhtBBrd).Activate
        rowuhtBBrd = ActiveCell.Offset(3, 0).Row
        
        Cells.Find(What:=stoBBrd).Activate
        rowstoBBrd = ActiveCell.Offset(24, 0).Row

        Cells.Find(What:=frozPBrd).Activate
        rowfrozPBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=conMBrd).Activate
        rowconMBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=valLBrd).Activate
        rowvalLBrd = ActiveCell.Row
        
        Cells.Find(What:=buttBrd).Activate
        rowbuttBrd = ActiveCell.Row

        Cells.Find(What:=fwmBrd).Activate
        rowfwmBrd = ActiveCell.Offset(5, 0).Row
       
        valCSFCh = Worksheets(wsCSFSh).Cells(cCSFRow, mthCOl)
        valcultBrd = Worksheets(wsCSFSh).Cells(rowcultBrd, mthCOl)
        valeverBrd = Worksheets(wsCSFSh).Cells(roweverBrd, mthCOl)
        valfresBBrd = Worksheets(wsCSFSh).Cells(rowfresBBrd, mthCOl)
        valspecCBrd = Worksheets(wsCSFSh).Cells(rowspecCBrd, mthCOl)
        valspreBrd = Worksheets(wsCSFSh).Cells(rowspreBrd, mthCOl)
        valyogDBrd = Worksheets(wsCSFSh).Cells(rowyogDBrd, mthCOl)
        valflouMBrd = Worksheets(wsCSFSh).Cells(rowflouMBrd, mthCOl)
        valoilDBrd = Worksheets(wsCSFSh).Cells(rowoilDBrd, mthCOl)
        valsweBBrd = Worksheets(wsCSFSh).Cells(rowsweBBrd, mthCOl)
        valuhtBBrd = Worksheets(wsCSFSh).Cells(rowuhtBBrd, mthCOl)
        valstoBBrd = Worksheets(wsCSFSh).Cells(rowstoBBrd, mthCOl)
        valfrozPBrd = Worksheets(wsCSFSh).Cells(rowfrozPBrd, mthCOl)
        valconMBrd = Worksheets(wsCSFSh).Cells(rowconMBrd, mthCOl)
        valvalLBrd = Worksheets(wsCSFSh).Cells(rowvalLBrd, mthCOl)
        valbuttBrd = Worksheets(wsCSFSh).Cells(rowbuttBrd, mthCOl)
        valfwmBrd = Worksheets(wsCSFSh).Cells(rowfwmBrd, mthCOl)
        totChl = Worksheets(wsCSFSh).Cells(totChlRow, mthCOl)
        totGro = Worksheets(wsCSFSh).Cells(totGroRow, mthCOl)
        totReb = Worksheets(wsCSFSh).Cells(totRebRow, mthCOl)
        totOth = Worksheets(wsCSFSh).Cells(othRebRow, mthCOl)
        totGrndReb = Worksheets(wsCSFSh).Cells(totGrdRow, mthCOl)
        
        Worksheets("csfSummary").Range("F" & k) = valCSFCh
        Worksheets("csfSummary").Range("H" & k) = totChl
        Worksheets("csfSummary").Range("I" & k) = totGro
        Worksheets("csfSummary").Range("J" & k) = totReb
        Worksheets("csfSummary").Range("K" & k) = totOth
        Worksheets("csfSummary").Range("L" & k) = totGrndReb
        Worksheets("csfSummary").Range("X" & k) = valcultBrd
        Worksheets("csfSummary").Range("Y" & k) = valeverBrd
        Worksheets("csfSummary").Range("Z" & k) = valfresBBrd
        Worksheets("csfSummary").Range("AA" & k) = valspecCBrd
        Worksheets("csfSummary").Range("AB" & k) = valspreBrd
        Worksheets("csfSummary").Range("AC" & k) = valyogDBrd
        Worksheets("csfSummary").Range("AD" & k) = valflouMBrd
        Worksheets("csfSummary").Range("AE" & k) = valoilDBrd
        Worksheets("csfSummary").Range("AF" & k) = valsweBBrd
        Worksheets("csfSummary").Range("AG" & k) = valuhtBBrd
        Worksheets("csfSummary").Range("AH" & k) = valstoBBrd
        Worksheets("csfSummary").Range("AI" & k) = valfrozPBrd
        Worksheets("csfSummary").Range("AJ" & k) = valconMBrd
        Worksheets("csfSummary").Range("AK" & k) = valvalLBrd
        Worksheets("csfSummary").Range("AL" & k) = valbuttBrd
        Worksheets("csfSummary").Range("AM" & k) = valfwmBrd
        
        ElseIf agmtType(i, 1) = "oneGF" And arrPayFreq(i, 1) = "Qtr" Then
        'bake
        
        ''''''''''''
        '''''CSF''''
        ''''''''''''
        
        Cells.Find(What:=bake).Activate
        ActiveCell.End(xlUp).Select
        ActiveCell.End(xlUp).Select
        ActiveCell.End(xlUp).Select
        bCSFRow = ActiveCell.Row
        Debug.Print bCSFRow
         
    '   Chilled
    '   Cells.Find(What:=chill).Activate
        Cells.Find(What:=bake).Activate
        ActiveCell.End(xlUp).Select
        cCSFRow = ActiveCell.Row
        
        Cells.Find(What:=mth).Activate
        mthCOl = ActiveCell.Column
        
        
        'subtotal
        Cells.Find(What:=bakeTotal).Activate
        totBakRow = ActiveCell.Row
        
        Cells.Find(What:=chillTotal).Activate
        totChlRow = ActiveCell.Row
        
        Cells.Find(What:=GrocTotal).Activate
        totGroRow = ActiveCell.Row
        
        Cells.Find(What:=othTotal).Activate
        othRebRow = ActiveCell.Row
        
        Cells.Find(What:=RebateTotal).Activate
        totRebRow = ActiveCell.Row
        
        Cells.Find(What:=grndTotal).Activate
        totGrdRow = ActiveCell.Row
        
        
        'brand
        
        Cells.Find(What:=homBrd).Activate
        rowhomBrd = ActiveCell.Offset(3, 0).Row

        Cells.Find(What:=loafBrd).Activate
        rowloafBrd = ActiveCell.Offset(7, 0).Row

        Cells.Find(What:=occBBrd).Activate
        rowoccBBrd = ActiveCell.Offset(5, 0).Row
        'chill
        Cells.Find(What:=cultBrd).Activate
        rowcultBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=everBrd).Activate
        roweverBrd = ActiveCell.Offset(2, 0).Row
        
        Cells.Find(What:=fresBBrd).Activate
        rowfresBBrd = ActiveCell.Offset(10, 0).Row

        Cells.Find(What:=specCBrd).Activate
        rowspecCBrd = ActiveCell.Offset(3, 0).Row

        Cells.Find(What:=spreBrd).Activate
        rowspreBrd = ActiveCell.Offset(5, 0).Row
        
        Cells.Find(What:=yogDBrd).Activate
        rowyogDBrd = ActiveCell.Offset(4, 0).Row

        Cells.Find(What:=flouMBrd).Activate
        rowflouMBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=oilDBrd).Activate
        rowoilDBrd = ActiveCell.Offset(5, 0).Row
        
        Cells.Find(What:=sweBBrd).Activate
        rowsweBBrd = ActiveCell.Offset(4, 0).Row

        Cells.Find(What:=uhtBBrd).Activate
        rowuhtBBrd = ActiveCell.Offset(3, 0).Row
        
        Cells.Find(What:=stoBBrd).Activate
        rowstoBBrd = ActiveCell.Offset(24, 0).Row

        Cells.Find(What:=frozPBrd).Activate
        rowfrozPBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=conMBrd).Activate
        rowconMBrd = ActiveCell.Offset(2, 0).Row

        Cells.Find(What:=valLBrd).Activate
        rowvalLBrd = ActiveCell.Row
        
        Cells.Find(What:=buttBrd).Activate
        rowbuttBrd = ActiveCell.Row

        Cells.Find(What:=fwmBrd).Activate
        rowfwmBrd = ActiveCell.Offset(5, 0).Row

        
        valCSFBk = Worksheets(wsCSFSh).Cells(bCSFRow, mthCOl)
        valCSFCh = Worksheets(wsCSFSh).Cells(cCSFRow, mthCOl)
        totBak = Worksheets(wsCSFSh).Cells(totBakRow, mthCOl)
        totChl = Worksheets(wsCSFSh).Cells(totChlRow, mthCOl)
        totGro = Worksheets(wsCSFSh).Cells(totGroRow, mthCOl)
        totOth = Worksheets(wsCSFSh).Cells(othRebRow, mthCOl)
        totReb = Worksheets(wsCSFSh).Cells(totRebRow, mthCOl)
        totGrndReb = Worksheets(wsCSFSh).Cells(totGrdRow, mthCOl)
        
        valhomBrd = Worksheets(wsCSFSh).Cells(rowhomBrd, mthCOl)
        valloafBrd = Worksheets(wsCSFSh).Cells(rowloafBrd, mthCOl)
        valoccBBrd = Worksheets(wsCSFSh).Cells(rowoccBBrd, mthCOl)
        valcultBrd = Worksheets(wsCSFSh).Cells(rowcultBrd, mthCOl)
        valeverBrd = Worksheets(wsCSFSh).Cells(roweverBrd, mthCOl)
        valfresBBrd = Worksheets(wsCSFSh).Cells(rowfresBBrd, mthCOl)
        valspecCBrd = Worksheets(wsCSFSh).Cells(rowspecCBrd, mthCOl)
        valspreBrd = Worksheets(wsCSFSh).Cells(rowspreBrd, mthCOl)
        valyogDBrd = Worksheets(wsCSFSh).Cells(rowyogDBrd, mthCOl)
        valflouMBrd = Worksheets(wsCSFSh).Cells(rowflouMBrd, mthCOl)
        valoilDBrd = Worksheets(wsCSFSh).Cells(rowoilDBrd, mthCOl)
        valsweBBrd = Worksheets(wsCSFSh).Cells(rowsweBBrd, mthCOl)
        valuhtBBrd = Worksheets(wsCSFSh).Cells(rowuhtBBrd, mthCOl)
        valstoBBrd = Worksheets(wsCSFSh).Cells(rowstoBBrd, mthCOl)
        valfrozPBrd = Worksheets(wsCSFSh).Cells(rowfrozPBrd, mthCOl)
        valconMBrd = Worksheets(wsCSFSh).Cells(rowconMBrd, mthCOl)
        valvalLBrd = Worksheets(wsCSFSh).Cells(rowvalLBrd, mthCOl)
        valbuttBrd = Worksheets(wsCSFSh).Cells(rowbuttBrd, mthCOl)
        valfwmBrd = Worksheets(wsCSFSh).Cells(rowfwmBrd, mthCOl)


        Worksheets("csfSummary").Range("E" & k) = valCSFBk
        Worksheets("csfSummary").Range("F" & k) = valCSFCh
        Worksheets("csfSummary").Range("G" & k) = totBak
        Worksheets("csfSummary").Range("H" & k) = totChl
        Worksheets("csfSummary").Range("I" & k) = totGro
        Worksheets("csfSummary").Range("J" & k) = totReb
        Worksheets("csfSummary").Range("K" & k) = totOth
        Worksheets("csfSummary").Range("L" & k) = totGrndReb
        
        Worksheets("csfSummary").Range("U" & k) = valhomBrd
        Worksheets("csfSummary").Range("V" & k) = valloafBrd
        Worksheets("csfSummary").Range("W" & k) = valoccBBrd
        Worksheets("csfSummary").Range("X" & k) = valcultBrd
        Worksheets("csfSummary").Range("Y" & k) = valeverBrd
        Worksheets("csfSummary").Range("Z" & k) = valfresBBrd
        Worksheets("csfSummary").Range("AA" & k) = valspecCBrd
        Worksheets("csfSummary").Range("AB" & k) = valspreBrd
        Worksheets("csfSummary").Range("AC" & k) = valyogDBrd
        Worksheets("csfSummary").Range("AD" & k) = valflouMBrd
        Worksheets("csfSummary").Range("AE" & k) = valoilDBrd
        Worksheets("csfSummary").Range("AF" & k) = valsweBBrd
        Worksheets("csfSummary").Range("AG" & k) = valuhtBBrd
        Worksheets("csfSummary").Range("AH" & k) = valstoBBrd
        Worksheets("csfSummary").Range("AI" & k) = valfrozPBrd
        Worksheets("csfSummary").Range("AJ" & k) = valconMBrd
        Worksheets("csfSummary").Range("AK" & k) = valvalLBrd
        Worksheets("csfSummary").Range("AL" & k) = valbuttBrd
        Worksheets("csfSummary").Range("AM" & k) = valfwmBrd

        
        
        
        ElseIf agmtType(i, 1) = "bake" And arrPayFreq(i, 1) = "Qtr" Then
        
        If ActiveSheet.Name = "NWThorndonP" Then GoTo handleStore:
        Cells.Find(What:=closBal).Activate
        bCSFRow = ActiveCell.Row
        


        Cells.Find(What:=mth).Activate
        mthCOl = ActiveCell.Column
handleStore:
        'brand
        Cells.Find(What:=loaBrdBake).Activate
        rowloaBrdBake = ActiveCell.Row

        Cells.Find(What:=homBrdBake).Activate
        rowhomBrdBake = ActiveCell.Row
        
        Cells.Find(What:=occBrdBake).Activate
        rowoccBrdBake = ActiveCell.Row
        

        'subtot
        Cells.Find(What:=bakeBrd).Activate
        totBakBRow = ActiveCell.Row
        
        valloaBrdBake = Worksheets(wsCSFSh).Cells(rowloaBrdBake, mthCOl)
        valhomBrdBake = Worksheets(wsCSFSh).Cells(rowhomBrdBake, mthCOl)
        valoccBrdBake = Worksheets(wsCSFSh).Cells(rowoccBrdBake, mthCOl)
        valCSFBk = Worksheets(wsCSFSh).Cells(bCSFRow, mthCOl)
        totBakB = Worksheets(wsCSFSh).Cells(totBakBRow, mthCOl)
        
        Worksheets("csfSummary").Range("E" & k) = valCSFBk
        Worksheets("csfSummary").Range("G" & k) = totBakB
        Worksheets("csfSummary").Range("U" & k) = valhomBrdBake
        Worksheets("csfSummary").Range("V" & k) = valloaBrdBake
        Worksheets("csfSummary").Range("W" & k) = valoccBrdBake

        
        End If
        
    End With
    
    i = i + 1

Next k


Worksheets("csfSummary").Activate

With ActiveSheet
    Columns.AutoFit
    Columns("E:AM").NumberFormat = "$#,###;[Red]($#,###);0"
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    .Range("M2").Formula = "=VLOOKUP(A2,bakingRebate,bakingRebateCol,0)*G2"
    .Range("N2").Formula = "=VLOOKUP(A2,nonWhiteClash,nonWhiteClashCol,0)*G2"
    .Range("O2").Formula = "=VLOOKUP(A2,chilledRebate,chilledRebateCol,0)*H2"
    .Range("P2").Formula = "=VLOOKUP(A2,yogSS,yogSSCol,0)*AC2"
    .Range("Q2").Formula = "=VLOOKUP(A2,cultSS,cultSSCol,0)*X2"
    .Range("R2").Formula = "=VLOOKUP(A2,groceryRebate,groceryRebateCol,0)*I2"
    .Range("S2").Formula = "=VLOOKUP(A2,oneGF,oneGFCol,FALSE)*L2"
    .Range("T2").Formula = "=M2+N2+O2+P2+Q2+R2"
    .Range("AO2").Formula = "=-G2+U2+V2+W2"
    .Range("AP2").Formula = "=-H2+X2+Y2+Z2+AA2+AB2+AC2"
    .Range("AQ2").Formula = "=-I2+AD2+AE2+AF2+AG2"
    .Range("AR2").Formula = "=-K2+AH2+AI2+AJ2+AK2+AL2+AM2"
    For i = 2 To lrow
    'agmtType(i, 1) = "oneGF"
    If .Range("B" & i) = "oneGF" Then
    .Range("AS" & i).Formula = "=-L" & i & "+U" & i & "+V" & i & "+W" & i & "+X" & i & "+Y" & i & "+Z" & i & "+AA" & i & "+AB" & i & "+AC" & i & "+AD" & i & "+AE" & i & "+AF" & i & "+AG" & i & "+AH" & i & "+AI" & i & "+AJ" & i & "+AK" & i & "+AL" & i & "+AM" & i
 '   .Range("AS" & i).Formula = "=-L2+U2+V2+W2+X2+Y2+Z2+AA2+AB2+AC2+AD2+AE2+AF2+AG2+AH2+AI2+AJ2+AK2+AL2+AM2"
    ElseIf .Range("B" & i) = "bake" Then
    .Range("AS" & i).Formula = "=-G" & i & "+U" & i & "+V" & i & "+W" & i
    
    End If
    .Calculate
    Next i
    .Range("M2:T" & lrow).FillDown
    .Range("AO2:AR" & lrow).FillDown
    
    '=VLOOKUP($B$4,groceryRebate,groceryRebateCol,FALSE)
    '=VLOOKUP($B$4,cultSS,cultSSCol,0)
    '=VLOOKUP($B$5,yogSS,yogSSCol,0)
    ' =VLOOKUP($B$5,chilledRebate,chilledRebateCol,0)
    ' =VLOOKUP($B$5,nonWhiteClash,nonWhiteClashCol,0)
'    Range("E" & lrow + 1).Formula = "=sum(E2:E" & lrow

    .Calculate
End With


bakeTemp = "bakeTemplate"

With ActiveSheet
    On Error Resume Next
    Cells.Find(bakeTemp).Activate
    ActiveCell.EntireRow.Delete
    Range("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    Range("U1:AM1").Select
    Selection.Columns.Group
    Range("AO1:AS1").Select
    Selection.Columns.Group
    
End With


End Sub
