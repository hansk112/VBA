Attribute VB_Name = "rmWage"
Sub rmWage()
Attribute rmWage.VB_ProcData.VB_Invoke_Func = " \n14"

Range("R73:AC73").ClearContents
Range("R76:AC81").ClearContents
Range("R83:AC84").ClearContents
Range("R94:AC96").ClearContents
Range("R100:AC100").ClearContents
Range("R108:AC111").ClearContents
Range("R113:AC114").ClearContents


End Sub

 
Sub REFRESH_DATA()
Attribute REFRESH_DATA.VB_ProcData.VB_Invoke_Func = " \n14"

Dim epm As New FPMXLClient.EPMAddInAutomation

epm.RefreshActiveWorkBook

End Sub

Sub SAVE_DATA()
Attribute SAVE_DATA.VB_ProcData.VB_Invoke_Func = " \n14"

Dim epm As New FPMXLClient.EPMAddInAutomation
epm.SaveAndRefreshWorksheetData

End Sub

Sub fillCar()

Range("S117:AD125") = 3
Range("S127:AD134") = 3

End Sub


Sub fCar()

Cells.Find("Car Size(1.Small/2.Medium/3.Large)").Activate
ActiveCell.Offset(0, 1).Activate
topcar = ActiveCell.Row
btmcar = ActiveCell.End(xlDown).Row
Dim bRow As String
bRow = " EPMLocalMember(""Blank Row"",""012"",""000"")"

For i = 1 To btmcar
   j = 1
    If ActiveCell <> "" Then
    ActiveCell.Offset(j, 0).Activate
    ElseIf ActiveCell.Formula = "= EPMLocalMember(""Blank Row"",""012"",""000"")" = bRow Then
    GoTo skiploop:
    End If
    'j = j + 1
Next i
skiploop:
btmcar = ActiveCell.Row

Debug.Print topcar & " " & btmcar

End Sub




Sub PasswordBreaker()

'Breaks worksheet password protection.

Dim i As Integer, j As Integer, k As Integer
Dim l As Integer, m As Integer, n As Integer
Dim i1 As Integer, i2 As Integer, i3 As Integer
Dim i4 As Integer, i5 As Integer, i6 As Integer
On Error Resume Next
For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
If ActiveSheet.ProtectContents = False Then
MsgBox "One usable password is " & Chr(i) & Chr(j) & _
Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
Exit Sub
End If
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub

Sub allocateSpreadBPC()
Dim nrows As Long, col As Long, tot As Long
Dim tot4 As Double, tot5 As Double, fourWk As Double, fiveWk As Double
Dim gl As String, glT As String
num = Range("AH55").Interior.Color
nrows = Cells(Rows.Count, 20).End(xlUp).Row

labCost = "BPC-LAB - Labour Costs"

Cells.Find(labCost).Activate
strtLoop = ActiveCell.Row
strtLoop = strtLoop + 1

For i = strtLoop To nrows
col = Range("Ah" & i).Interior.Color
col2 = Range("S" & i).Interior.Color
gl = Range("Q" & i)



pcard = "GL68963 - Purchase Card Trxs"

fourWk = 0.0769
fiveWk = 0.0961


Debug.Print gl
    'grey color cell
    If col = 10855845 Or gl = pcard Then GoTo skipLine:
    
    If col = 16777215 Then
    tot4 = Range("AF" & i)
    tot5 = Range("AF" & i)
    
    fuel = "GL64105 - Vehicles Fuel"
    rego = "GL64110 - Vehicles Rego"
    serv = "GL64115 - Vehicles Service"
    rent = "GL64125 - Vehicles Rent"
    
    merchGL = "GL61460 - Merchandising"

    tot4 = (tot4 * fourWk)
    tot5 = (tot5 * fiveWk)
    
    If gl = fuel Or gl = rego Or gl = serv Or gl = rent Then
    tot4 = (tot4 * 2)
    tot5a = (tot5 * 2)
    End If
    
'    If gl = merchGL Then
'        tot4 = tot4 * 1.1
'        tot5 = tot5 * 1.1
'    End If
    
    Range("S" & i) = tot4
    Range("T" & i) = tot4
    Range("U" & i) = tot5
    Range("V" & i) = tot4
    Range("W" & i) = tot4
    Range("X" & i) = tot5
    Range("Y" & i) = tot4
    Range("Z" & i) = tot4
    Range("AA" & i) = tot5
    Range("AB" & i) = tot4
    Range("AC" & i) = tot4
    Range("AD" & i) = tot5
    
    End If
'    Range("AF" & i).Formula = "=" & tot "/12"
    
'  Debug.Print col = Range("S" & i)

skipLine:
Next i

Debug.Print num

MsgBox "done"
End Sub


Function trimGL(glCode As String) As String
glCode = Trim(Left(glCode, 7))


End Function

Sub travelDisagg()

'28812
Dim nrows As Long

Dim incr As Double
incr = InputBox("Enter Travel For Yr", "Disagg Travel")

incr = incr / 12

nrows = Cells(Rows.Count, 17).End(xlUp).Row

srt = "BPC-MER - Merchandising Costs"
trv = "BPC-TRAV - Travel"

Cells.Find(srt).Activate
numStart = ActiveCell.Row
numStart = numStart + 1


Cells.Find(trv).Activate
JanTot = ActiveCell.Offset(0, 2)
FebTot = ActiveCell.Offset(0, 3)
MarTot = ActiveCell.Offset(0, 4)
AprTot = ActiveCell.Offset(0, 5)
MayTot = ActiveCell.Offset(0, 6)
JunTot = ActiveCell.Offset(0, 7)
JulTot = ActiveCell.Offset(0, 8)
AugTot = ActiveCell.Offset(0, 9)
SepTot = ActiveCell.Offset(0, 10)
OctTot = ActiveCell.Offset(0, 11)
NovTot = ActiveCell.Offset(0, 12)
DecTot = ActiveCell.Offset(0, 13)



For i = numStart To nrows

If Range("Q" & i) = trv Then
    Exit Sub
End If

   JanGL = Range("S" & i)
   FebGl = Range("T" & i)
   MarGL = Range("U" & i)
   AprGL = Range("V" & i)
   MayGL = Range("W" & i)
   JunGL = Range("X" & i)
   JulGL = Range("Y" & i)
   AugGl = Range("Z" & i)
   SepGl = Range("AA" & i)
   OctGL = Range("AB" & i)
   NovGL = Range("AC" & i)
   DecGL = Range("AD" & i)
   
   JandAgg = JanGL / JanTot * incr
   FebdAgg = FebGl / FebTot * incr
   MardAgg = MarGL / MarTot * incr
   AprdAgg = AprGL / AprTot * incr
   MaydAgg = MayGL / MayTot * incr
   JundAgg = JunGL / JunTot * incr
   JuldAgg = JulGL / JulTot * incr
   AugdAgg = AugGl / AugTot * incr
   SepJdAgg = SepGl / SepTot * incr
   OctJdAgg = OctGL / OctTot * incr
   NovJdAgg = NovGL / NovTot * incr
   DecJdAgg = DecGL / DecTot * incr
   
      Range("S" & i) = JandAgg + JanGL
      Range("T" & i) = FebdAgg + FebGl
      Range("U" & i) = MardAgg + MarGL
      Range("V" & i) = AprdAgg + AprGL
      Range("W" & i) = MaydAgg + MayGL
      Range("X" & i) = JundAgg + JunGL
      Range("Y" & i) = JuldAgg + JulGL
      Range("Z" & i) = AugdAgg + AugGl
      Range("AA" & i) = SepJdAgg + SepGl
      Range("AB" & i) = OctJdAgg + OctGL
      Range("AC" & i) = NovJdAgg + NovGL
      Range("AD" & i) = DecJdAgg + DecGL


Debug.Print JandAgg & FebdAgg

Next i

MsgBox done
End Sub

Sub MerchDisagg()

Dim nrows As Long

Dim incr As Double
incr = InputBox("Enter Travel For Yr", "Disagg Travel")

incr = incr

nrows = Cells(Rows.Count, 17).End(xlUp).Row

srt = "BPC-MKT - Marketing Costs"
trv = "BPC-MER - Merchandising Costs"
Lab = "BPC-LAB - Labour Costs"
mops = "BPC -OPS - Mfg & Operations"
FDis "BPC-FRE - Freight & Dist."
War = "BPC -WAR - Warehousing"
MarCos = "BPC-MKT - Marketing Costs"
MerCos = "BPC-MER - Merchandising Costs"
TravCos = "BPC -TRAV - Travel"
MotoV = "BPC-MVEXP - Motor Vehicle Expens"
Prof = "BPC-PFEE - Professional Fees"
Insu = "BPC -INS - Insurance"
Train = "BPC -TRAN - Training"
Conf = "BPC -CONF - Conferences"
OccuCos = "BPC-RATE - Occupany Costs"
ITCos = "BPC -IT - IT"
Comm = "BPC-COMMS - Communication Costs"
RandD = "BPC-RND - R&D"
RepAndMai = "BPC -REP - Repairs & Maintenana"


    
    

Cells.Find(srt).Activate
numStart = ActiveCell.Row
numStart = numStart + 1


Cells.Find(trv).Activate
JanTot = ActiveCell.Offset(0, 2)
FebTot = ActiveCell.Offset(0, 3)
MarTot = ActiveCell.Offset(0, 4)
AprTot = ActiveCell.Offset(0, 5)
MayTot = ActiveCell.Offset(0, 6)
JunTot = ActiveCell.Offset(0, 7)
JulTot = ActiveCell.Offset(0, 8)
AugTot = ActiveCell.Offset(0, 9)
SepTot = ActiveCell.Offset(0, 10)
OctTot = ActiveCell.Offset(0, 11)
NovTot = ActiveCell.Offset(0, 12)
DecTot = ActiveCell.Offset(0, 13)



For i = numStart To nrows

If Range("Q" & i) = mops Then
    Exit Sub
End If

   JanGL = Range("S" & i)
   FebGl = Range("T" & i)
   MarGL = Range("U" & i)
   AprGL = Range("V" & i)
   MayGL = Range("W" & i)
   JunGL = Range("X" & i)
   JulGL = Range("Y" & i)
   AugGl = Range("Z" & i)
   SepGl = Range("AA" & i)
   OctGL = Range("AB" & i)
   NovGL = Range("AC" & i)
   DecGL = Range("AD" & i)
   
   JandAgg = JanGL / JanTot * incr
   FebdAgg = FebGl / FebTot * incr
   MardAgg = MarGL / MarTot * incr
   AprdAgg = AprGL / AprTot * incr
   MaydAgg = MayGL / MayTot * incr
   JundAgg = JunGL / JunTot * incr
   JuldAgg = JulGL / JulTot * incr
   AugdAgg = AugGl / AugTot * incr
   SepJdAgg = SepGl / SepTot * incr
   OctJdAgg = OctGL / OctTot * incr
   NovJdAgg = NovGL / NovTot * incr
   DecJdAgg = DecGL / DecTot * incr
   
      Range("S" & i) = JandAgg + JanGL
      Range("T" & i) = FebdAgg + FebGl
      Range("U" & i) = MardAgg + MarGL
      Range("V" & i) = AprdAgg + AprGL
      Range("W" & i) = MaydAgg + MayGL
      Range("X" & i) = JundAgg + JunGL
      Range("Y" & i) = JuldAgg + JulGL
      Range("Z" & i) = AugdAgg + AugGl
      Range("AA" & i) = SepJdAgg + SepGl
      Range("AB" & i) = OctJdAgg + OctGL
      Range("AC" & i) = NovJdAgg + NovGL
      Range("AD" & i) = DecJdAgg + DecGL


Debug.Print JandAgg & FebdAgg

Next i

End Sub


Sub MktOpsDisagg()
'11161
Dim nrows As Long

Dim incr As Double
incr = InputBox("Enter Travel For Yr", "Disagg Travel")

incr = incr / 12

nrows = Cells(Rows.Count, 17).End(xlUp).Row

srt = "BPC-MKT - Marketing Costs"
trv = "BPC-MER - Merchandising Costs"
Lab = "BPC-LAB - Labour Costs"
mops = "BPC-OPS - Mfg & Operations"
FDis = "BPC-FRE - Freight & Dist."
War = "BPC -WAR - Warehousing"
MarCos = "BPC-MKT - Marketing Costs"
MerCos = "BPC-MER - Merchandising Costs"
TravCos = "BPC -TRAV - Travel"
MotoV = "BPC-MVEXP - Motor Vehicle Expens"
Prof = "BPC-PFEE - Professional Fees"
Insu = "BPC -INS - Insurance"
Train = "BPC -TRAN - Training"
Conf = "BPC -CONF - Conferences"
OccuCos = "BPC-RATE - Occupany Costs"
ITCos = "BPC -IT - IT"
Comm = "BPC-COMMS - Communication Costs"
RandD = "BPC-RND - R&D"
RepAndMai = "BPC -REP - Repairs & Maintenana"


    
    

Cells.Find(Lab).Activate
numStart = ActiveCell.Row
numStart = numStart + 1


Cells.Find(mops).Activate
JanTot = ActiveCell.Offset(0, 2)
FebTot = ActiveCell.Offset(0, 3)
MarTot = ActiveCell.Offset(0, 4)
AprTot = ActiveCell.Offset(0, 5)
MayTot = ActiveCell.Offset(0, 6)
JunTot = ActiveCell.Offset(0, 7)
JulTot = ActiveCell.Offset(0, 8)
AugTot = ActiveCell.Offset(0, 9)
SepTot = ActiveCell.Offset(0, 10)
OctTot = ActiveCell.Offset(0, 11)
NovTot = ActiveCell.Offset(0, 12)
DecTot = ActiveCell.Offset(0, 13)



For i = numStart To nrows

If Range("Q" & i) = mops Then
    Exit Sub
End If

   JanGL = Range("S" & i)
   FebGl = Range("T" & i)
   MarGL = Range("U" & i)
   AprGL = Range("V" & i)
   MayGL = Range("W" & i)
   JunGL = Range("X" & i)
   JulGL = Range("Y" & i)
   AugGl = Range("Z" & i)
   SepGl = Range("AA" & i)
   OctGL = Range("AB" & i)
   NovGL = Range("AC" & i)
   DecGL = Range("AD" & i)
   
   JandAgg = JanGL / JanTot * incr
   FebdAgg = FebGl / FebTot * incr
   MardAgg = MarGL / MarTot * incr
   AprdAgg = AprGL / AprTot * incr
   MaydAgg = MayGL / MayTot * incr
   JundAgg = JunGL / JunTot * incr
   JuldAgg = JulGL / JulTot * incr
   AugdAgg = AugGl / AugTot * incr
   SepJdAgg = SepGl / SepTot * incr
   OctJdAgg = OctGL / OctTot * incr
   NovJdAgg = NovGL / NovTot * incr
   DecJdAgg = DecGL / DecTot * incr
   
      Range("S" & i) = JandAgg + JanGL
      Range("T" & i) = FebdAgg + FebGl
      Range("U" & i) = MardAgg + MarGL
      Range("V" & i) = AprdAgg + AprGL
      Range("W" & i) = MaydAgg + MayGL
      Range("X" & i) = JundAgg + JunGL
      Range("Y" & i) = JuldAgg + JulGL
      Range("Z" & i) = AugdAgg + AugGl
      Range("AA" & i) = SepJdAgg + SepGl
      Range("AB" & i) = OctJdAgg + OctGL
      Range("AC" & i) = NovJdAgg + NovGL
      Range("AD" & i) = DecJdAgg + DecGL


Debug.Print JandAgg & FebdAgg

Next i

End Sub

Sub ResandDevDisagg()
' 43752
Dim nrows As Long

Dim incr As Double
incr = InputBox("Enter R&D For Yr", "Disagg Travel")

incr = incr / 12

nrows = Cells(Rows.Count, 17).End(xlUp).Row

srt = "BPC-MKT - Marketing Costs"
trv = "BPC-MER - Merchandising Costs"
Lab = "BPC-LAB - Labour Costs"
mops = "BPC-OPS - Mfg & Operations"
FDis = "BPC-FRE - Freight & Dist."
War = "BPC -WAR - Warehousing"
MarCos = "BPC-MKT - Marketing Costs"
MerCos = "BPC-MER - Merchandising Costs"
TravCos = "BPC -TRAV - Travel"
MotoV = "BPC-MVEXP - Motor Vehicle Expens"
Prof = "BPC-PFEE - Professional Fees"
Insu = "BPC -INS - Insurance"
Train = "BPC -TRAN - Training"
Conf = "BPC -CONF - Conferences"
OccuCos = "BPC-RATE - Occupany Costs"
ITCos = "BPC -IT - IT"
Comm = "BPC-COMMS - Communication Costs"
RD = "BPC-RND - R&D"
RepAndMai = "BPC -REP - Repairs & Maintenana"


    
    

Cells.Find(Comm).Activate
numStart = ActiveCell.Row
numStart = numStart + 1


Cells.Find(RD).Activate
JanTot = ActiveCell.Offset(0, 2)
FebTot = ActiveCell.Offset(0, 3)
MarTot = ActiveCell.Offset(0, 4)
AprTot = ActiveCell.Offset(0, 5)
MayTot = ActiveCell.Offset(0, 6)
JunTot = ActiveCell.Offset(0, 7)
JulTot = ActiveCell.Offset(0, 8)
AugTot = ActiveCell.Offset(0, 9)
SepTot = ActiveCell.Offset(0, 10)
OctTot = ActiveCell.Offset(0, 11)
NovTot = ActiveCell.Offset(0, 12)
DecTot = ActiveCell.Offset(0, 13)



For i = numStart To nrows

If Range("Q" & i) = RD Then
    Exit Sub
End If

   JanGL = Range("S" & i)
   FebGl = Range("T" & i)
   MarGL = Range("U" & i)
   AprGL = Range("V" & i)
   MayGL = Range("W" & i)
   JunGL = Range("X" & i)
   JulGL = Range("Y" & i)
   AugGl = Range("Z" & i)
   SepGl = Range("AA" & i)
   OctGL = Range("AB" & i)
   NovGL = Range("AC" & i)
   DecGL = Range("AD" & i)
   
   JandAgg = JanGL / JanTot * incr
   FebdAgg = FebGl / FebTot * incr
   MardAgg = MarGL / MarTot * incr
   AprdAgg = AprGL / AprTot * incr
   MaydAgg = MayGL / MayTot * incr
   JundAgg = JunGL / JunTot * incr
   JuldAgg = JulGL / JulTot * incr
   AugdAgg = AugGl / AugTot * incr
   SepJdAgg = SepGl / SepTot * incr
   OctJdAgg = OctGL / OctTot * incr
   NovJdAgg = NovGL / NovTot * incr
   DecJdAgg = DecGL / DecTot * incr
   
      Range("S" & i) = JandAgg + JanGL
      Range("T" & i) = FebdAgg + FebGl
      Range("U" & i) = MardAgg + MarGL
      Range("V" & i) = AprdAgg + AprGL
      Range("W" & i) = MaydAgg + MayGL
      Range("X" & i) = JundAgg + JunGL
      Range("Y" & i) = JuldAgg + JulGL
      Range("Z" & i) = AugdAgg + AugGl
      Range("AA" & i) = SepJdAgg + SepGl
      Range("AB" & i) = OctJdAgg + OctGL
      Range("AC" & i) = NovJdAgg + NovGL
      Range("AD" & i) = DecJdAgg + DecGL


Debug.Print JandAgg & FebdAgg

Next i

End Sub


Sub OtherDisagg()
' 43752
Dim nrows As Long

Dim incr As Double
incr = InputBox("Enter R&D For Yr", "Disagg Travel")

incr = incr / 12

nrows = Cells(Rows.Count, 17).End(xlUp).Row

srt = "BPC-MKT - Marketing Costs"
trv = "BPC-MER - Merchandising Costs"
Lab = "BPC-LAB - Labour Costs"
mops = "BPC-OPS - Mfg & Operations"
FDis = "BPC-FRE - Freight & Dist."
War = "BPC -WAR - Warehousing"
MarCos = "BPC-MKT - Marketing Costs"
MerCos = "BPC-MER - Merchandising Costs"
TravCos = "BPC -TRAV - Travel"
MotoV = "BPC-MVEXP - Motor Vehicle Expens"
Prof = "BPC-PFEE - Professional Fees"
Insu = "BPC -INS - Insurance"
Train = "BPC -TRAN - Training"
Conf = "BPC -CONF - Conferences"
OccuCos = "BPC-RATE - Occupany Costs"
ITCos = "BPC -IT - IT"
Comm = "BPC-COMMS - Communication Costs"
RD = "BPC-RND - R&D"
RepAndMai = "BPC -REP - Repairs & Maintenana"

oth = "BPC-OTH - Other expenses"

    
    

Cells.Find(RD).Activate
numStart = ActiveCell.Row
numStart = numStart + 1


Cells.Find(oth).Activate
JanTot = ActiveCell.Offset(0, 2)
FebTot = ActiveCell.Offset(0, 3)
MarTot = ActiveCell.Offset(0, 4)
AprTot = ActiveCell.Offset(0, 5)
MayTot = ActiveCell.Offset(0, 6)
JunTot = ActiveCell.Offset(0, 7)
JulTot = ActiveCell.Offset(0, 8)
AugTot = ActiveCell.Offset(0, 9)
SepTot = ActiveCell.Offset(0, 10)
OctTot = ActiveCell.Offset(0, 11)
NovTot = ActiveCell.Offset(0, 12)
DecTot = ActiveCell.Offset(0, 13)



For i = numStart To nrows

If Range("Q" & i) = oth Then
    Exit Sub
End If

   JanGL = Range("S" & i)
   FebGl = Range("T" & i)
   MarGL = Range("U" & i)
   AprGL = Range("V" & i)
   MayGL = Range("W" & i)
   JunGL = Range("X" & i)
   JulGL = Range("Y" & i)
   AugGl = Range("Z" & i)
   SepGl = Range("AA" & i)
   OctGL = Range("AB" & i)
   NovGL = Range("AC" & i)
   DecGL = Range("AD" & i)
   
   JandAgg = JanGL / JanTot * incr
   FebdAgg = FebGl / FebTot * incr
   MardAgg = MarGL / MarTot * incr
   AprdAgg = AprGL / AprTot * incr
   MaydAgg = MayGL / MayTot * incr
   JundAgg = JunGL / JunTot * incr
   JuldAgg = JulGL / JulTot * incr
   AugdAgg = AugGl / AugTot * incr
   SepJdAgg = SepGl / SepTot * incr
   OctJdAgg = OctGL / OctTot * incr
   NovJdAgg = NovGL / NovTot * incr
   DecJdAgg = DecGL / DecTot * incr
   
      Range("S" & i) = JandAgg + JanGL
      Range("T" & i) = FebdAgg + FebGl
      Range("U" & i) = MardAgg + MarGL
      Range("V" & i) = AprdAgg + AprGL
      Range("W" & i) = MaydAgg + MayGL
      Range("X" & i) = JundAgg + JunGL
      Range("Y" & i) = JuldAgg + JulGL
      Range("Z" & i) = AugdAgg + AugGl
      Range("AA" & i) = SepJdAgg + SepGl
      Range("AB" & i) = OctJdAgg + OctGL
      Range("AC" & i) = NovJdAgg + NovGL
      Range("AD" & i) = DecJdAgg + DecGL


Debug.Print JandAgg & FebdAgg

Next i

End Sub

Sub dynamicDisagg()
' 43752
Dim nrows As Long

Dim incr As Double


topGL = InputBox("Enter Top GL Node" & vbCr & _
"1: BPC-LAB - Labour Costs" & vbCr & _
"2: BPC-OPS - Mfg & Operations" & vbCr & _
"3: BPC-FRE - Freight & Dist." & vbCr & _
"4: BPC -WAR - Warehousing" & vbCr & _
"5: BPC-MKT - Marketing Costs" & vbCr & _
"6: BPC-MER - Merchandising Costs" & vbCr & _
"7: BPC -TRAV - Travel" & vbCr & _
"8: BPC-MVEXP - Motor Vehicle Expens" & vbCr & _
"9: BPC-PFEE - Professional Fees" & vbCr & _
"10: BPC -INS - Insurance" & vbCr & _
"11: BPC -TRAN - Training" & vbCr & _
"12: BPC -CONF - Conferences" & vbCr & _
"13: BPC-RATE - Occupany Costs" & vbCr & _
"14: BPC -IT - IT" & vbCr & _
"15: BPC-COMMS - Communication Costs" & vbCr & _
"16: BPC-RND - R&D" & vbCr & _
"17: BPC -REP - Repairs & Maintenana" & vbCr & _
"18: BPC-OTH - Other expenses", "Disagg Dynmaic")


Select Case topGL

Case 1
topGL = "BPC-LAB - Labour Costs"
Case 2
topGL = "BPC-OPS - Mfg & Operations"
Case 3
topGL = "BPC-FRE - Freight & Dist."
Case 4
topGL = "BPC -WAR - Warehousing"
Case 5
topGL = "BPC-MKT - Marketing Costs"
Case 6
topGL = "BPC-MER - Merchandising Costs"
Case 7
topGL = "BPC-TRAV - Travel"
Case 8
topGL = "BPC-MVEXP - Motor Vehicle Expens"
Case 9
topGL = "BPC-PFEE - Professional Fees"
Case 10
topGL = "BPC-INS - Insurance"
Case 11
topGL = "BPC-TRAN - Training"
Case 12
topGL = "BPC-CONF - Conferences"
Case 13
topGL = "BPC-RATE - Occupany Costs"
Case 14
topGL = "BPC-IT - IT"
Case 15
topGL = "BPC-COMMS - Communication Costs"
Case 16
topGL = "BPC-RND - R&D"
Case 17
topGL = "BPC-REP - Repairs & Maintenana"
Case 18
topGL = "BPC-OTH - Other expenses"

End Select

lowGL = InputBox("Enter Lower GL Node" & vbCr & _
"1: BPC-LAB - Labour Costs" & vbCr & _
"2: BPC-OPS - Mfg & Operations" & vbCr & _
"3: BPC-FRE - Freight & Dist." & vbCr & _
"4: BPC -WAR - Warehousing" & vbCr & _
"5: BPC-MKT - Marketing Costs" & vbCr & _
"6: BPC-MER - Merchandising Costs" & vbCr & _
"7: BPC -TRAV - Travel" & vbCr & _
"8: BPC-MVEXP - Motor Vehicle Expens" & vbCr & _
"9: BPC-PFEE - Professional Fees" & vbCr & _
"10: BPC -INS - Insurance" & vbCr & _
"11: BPC -TRAN - Training" & vbCr & _
"12: BPC -CONF - Conferences" & vbCr & _
"13: BPC-RATE - Occupany Costs" & vbCr & _
"14: BPC -IT - IT" & vbCr & _
"15: BPC-COMMS - Communication Costs" & vbCr & _
"16: BPC-RND - R&D" & vbCr & _
"17: BPC -REP - Repairs & Maintenana" & vbCr & _
"18: BPC-OTH - Other expenses", "Disagg Dynmaic")


Select Case lowGL

Case 1
lowGL = "BPC-LAB - Labour Costs"
Case 2
lowGL = "BPC-OPS - Mfg & Operations"
Case 3
lowGL = "BPC-FRE - Freight & Dist."
Case 4
lowGL = "BPC -WAR - Warehousing"
Case 5
lowGL = "BPC-MKT - Marketing Costs"
Case 6
lowGL = "BPC-MER - Merchandising Costs"
Case 7
lowGL = "BPC-TRAV - Travel"
Case 8
lowGL = "BPC-MVEXP - Motor Vehicle Expens"
Case 9
lowGL = "BPC-PFEE - Professional Fees"
Case 10
lowGL = "BPC-INS - Insurance"
Case 11
lowGL = "BPC-TRAN - Training"
Case 12
lowGL = "BPC-CONF - Conferences"
Case 13
lowGL = "BPC-RATE - Occupany Costs"
Case 14
lowGL = "BPC-IT - IT"
Case 15
lowGL = "BPC-COMMS - Communication Costs"
Case 16
lowGL = "BPC-RND - R&D"
Case 17
lowGL = "BPC-REP - Repairs & Maintenana"
Case 18
lowGL = "BPC-OTH - Other expenses"

End Select


incr = InputBox("Enter Disagg Amount For Year", "Disagg")
incr = incr / 12

nrows = Cells(Rows.Count, 17).End(xlUp).Row

srt = "BPC-MKT - Marketing Costs"
trv = "BPC-MER - Merchandising Costs"
Lab = "BPC-LAB - Labour Costs"
mops = "BPC-OPS - Mfg & Operations"
FDis = "BPC-FRE - Freight & Dist."
War = "BPC -WAR - Warehousing"
MarCos = "BPC-MKT - Marketing Costs"
MerCos = "BPC-MER - Merchandising Costs"
TravCos = "BPC -TRAV - Travel"
MotoV = "BPC-MVEXP - Motor Vehicle Expens"
Prof = "BPC-PFEE - Professional Fees"
Insu = "BPC -INS - Insurance"
Train = "BPC -TRAN - Training"
Conf = "BPC -CONF - Conferences"
OccuCos = "BPC-RATE - Occupany Costs"
ITCos = "BPC -IT - IT"
Comm = "BPC-COMMS - Communication Costs"
RD = "BPC-RND - R&D"
RepAndMai = "BPC -REP - Repairs & Maintenana"
oth = "BPC-OTH - Other expenses"

    
    

Cells.Find(topGL).Activate
numStart = ActiveCell.Row
numStart = numStart + 1
    

Cells.Find(lowGL).Activate
JanTot = ActiveCell.Offset(0, 2)
FebTot = ActiveCell.Offset(0, 3)
MarTot = ActiveCell.Offset(0, 4)
AprTot = ActiveCell.Offset(0, 5)
MayTot = ActiveCell.Offset(0, 6)
JunTot = ActiveCell.Offset(0, 7)
JulTot = ActiveCell.Offset(0, 8)
AugTot = ActiveCell.Offset(0, 9)
SepTot = ActiveCell.Offset(0, 10)
OctTot = ActiveCell.Offset(0, 11)
NovTot = ActiveCell.Offset(0, 12)
DecTot = ActiveCell.Offset(0, 13)



For i = numStart To nrows

If Range("Q" & i) = lowGL Then
    Exit Sub
End If

Set rng = Range("Q" & i)

If rng = srt Then
GoTo skipGL:
End If

Debug.Print rng
   JanGL = Range("S" & i)
   FebGl = Range("T" & i)
   MarGL = Range("U" & i)
   AprGL = Range("V" & i)
   MayGL = Range("W" & i)
   JunGL = Range("X" & i)
   JulGL = Range("Y" & i)
   AugGl = Range("Z" & i)
   SepGl = Range("AA" & i)
   OctGL = Range("AB" & i)
   NovGL = Range("AC" & i)
   DecGL = Range("AD" & i)
   
   JandAgg = JanGL / JanTot * incr
   FebdAgg = FebGl / FebTot * incr
   MardAgg = MarGL / MarTot * incr
   AprdAgg = AprGL / AprTot * incr
   MaydAgg = MayGL / MayTot * incr
   JundAgg = JunGL / JunTot * incr
   JuldAgg = JulGL / JulTot * incr
   AugdAgg = AugGl / AugTot * incr
   SepJdAgg = SepGl / SepTot * incr
   OctJdAgg = OctGL / OctTot * incr
   NovJdAgg = NovGL / NovTot * incr
   DecJdAgg = DecGL / DecTot * incr
   
      Range("S" & i) = JandAgg + JanGL
      Range("T" & i) = FebdAgg + FebGl
      Range("U" & i) = MardAgg + MarGL
      Range("V" & i) = AprdAgg + AprGL
      Range("W" & i) = MaydAgg + MayGL
      Range("X" & i) = JundAgg + JunGL
      Range("Y" & i) = JuldAgg + JulGL
      Range("Z" & i) = AugdAgg + AugGl
      Range("AA" & i) = SepJdAgg + SepGl
      Range("AB" & i) = OctJdAgg + OctGL
      Range("AC" & i) = NovJdAgg + NovGL
      Range("AD" & i) = DecJdAgg + DecGL


Debug.Print JandAgg & FebdAgg


skipGL:
Next i

End Sub


Sub allocateSpreadBPCBudget()
Dim nrows As Long, col As Long, tot As Long
Dim tot4 As Double, tot5 As Double, fourWk As Double, fiveWk As Double
Dim gl As String, glT As String
num = Range("AH55").Interior.Color
nrows = Cells(Rows.Count, 20).End(xlUp).Row

labCost = "BPC-LAB - Labour Costs"

Cells.Find(labCost).Activate
strtLoop = ActiveCell.Row
strtLoop = strtLoop + 1

For i = strtLoop To nrows
col = Range("Ah" & i).Interior.Color
col2 = Range("S" & i).Interior.Color
gl = Range("Q" & i)



pcard = "GL68963 - Purchase Card Trxs"

fourWk = 0.0769
fiveWk = 0.0961


Debug.Print gl
    'grey color cell
    If col = 10855845 Or gl = pcard Then GoTo skipLine:
    
    If col = 16777215 Then
    tot4 = Range("AE" & i)
    tot5 = Range("AE" & i)
    
    fuel = "GL64105 - Vehicles Fuel"
    rego = "GL64110 - Vehicles Rego"
    serv = "GL64115 - Vehicles Service"
    rent = "GL64125 - Vehicles Rent"
    
    tot4 = (tot4 * fourWk)
    tot5 = (tot5 * fiveWk)
    
    If gl = fuel Or gl = rego Or gl = serv Or gl = rent Then
    tot4 = (tot4 * 2)
    tot5a = (tot5 * 2)
    End If

    
    Range("S" & i) = tot4
    Range("T" & i) = tot4
    Range("U" & i) = tot5
    Range("V" & i) = tot4
    Range("W" & i) = tot4
    Range("X" & i) = tot5
    Range("Y" & i) = tot4
    Range("Z" & i) = tot4
    Range("AA" & i) = tot5
    Range("AB" & i) = tot4
    Range("AC" & i) = tot4
    Range("AD" & i) = tot5
    
    End If
'    Range("AF" & i).Formula = "=" & tot "/12"
    
'  Debug.Print col = Range("S" & i)

skipLine:
Next i

Debug.Print num

MsgBox "done"
End Sub

