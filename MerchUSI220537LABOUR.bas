Attribute VB_Name = "MerchuSI220537LABOUR"
Sub GSk()

fourWk = 0.0769
fiveWk = 0.0961

'wage ord
deductAnn = 200188 / 12


JanWO = 71473
FebWO = 81528
MarWO = 97362
AprWO = 81631
MayWO = 83786
JunWO = 1028531
JulWO = 83049
AugWO = 84011
SepWO = 99462
OctWO = 82342
NovWO = 79452
DecWO = 103142
'o/t
JanOT = 27543
FebOT = 19215
MarOT = 8940
AprOT = 13943
MayOT = 12797
JunOT = 13586
JulOT = 10160
AugOT = 14661
SepOT = 12850
OctOT = 11907
NovOT = 12331
DecOT = 12100

JanAL = 27543
FebAL = 19215
MarAL = 8940
AprAL = 13943
MayAL = 12797
JunAL = 13586
JulAL = 10160
AugAL = 14661
SepAL = 12850
OctAL = 11907
NovAL = 12331
DecAL = 12100



Range("R116").GoalSeek Goal:=JanWO, ChangingCell:=Range("R81") 'apr
Range("S116").GoalSeek Goal:=FebWO, ChangingCell:=Range("S81") 'may
Range("T116").GoalSeek Goal:=MarWO, ChangingCell:=Range("T81") 'jun
Range("U116").GoalSeek Goal:=AprWO, ChangingCell:=Range("U81")
Range("V116").GoalSeek Goal:=MayWO, ChangingCell:=Range("V81")
Range("W116").GoalSeek Goal:=JunWO, ChangingCell:=Range("W81")
Range("X116").GoalSeek Goal:=JulWO, ChangingCell:=Range("X81")
Range("Y116").GoalSeek Goal:=AugWO, ChangingCell:=Range("Y81")
Range("Z116").GoalSeek Goal:=SepWO, ChangingCell:=Range("Z81")
Range("AA116").GoalSeek Goal:=OctWO, ChangingCell:=Range("AA81")
Range("AB116").GoalSeek Goal:=NovWO, ChangingCell:=Range("AB81")
Range("AC116").GoalSeek Goal:=DecWO, ChangingCell:=Range("AC81")


'o/t
'Range("R120").GoalSeek Goal:=JanOT, ChangingCell:=Range("R80") 'apr
'Range("S120").GoalSeek Goal:=FebOT, ChangingCell:=Range("S80") 'may
'Range("T120").GoalSeek Goal:=MarOT, ChangingCell:=Range("T80") 'jun
'Range("U120").GoalSeek Goal:=AprOT, ChangingCell:=Range("U80")
'Range("V120").GoalSeek Goal:=MayOT, ChangingCell:=Range("V80")
'Range("W120").GoalSeek Goal:=JunOT, ChangingCell:=Range("W80")
'Range("X120").GoalSeek Goal:=JulOT, ChangingCell:=Range("X80")
'Range("Y120").GoalSeek Goal:=AugOT, ChangingCell:=Range("Y80")
'Range("Z120").GoalSeek Goal:=SepOT, ChangingCell:=Range("Z80")
'Range("AA120").GoalSeek Goal:=OctOT, ChangingCell:=Range("AA80")
'Range("AB120").GoalSeek Goal:=NovOT, ChangingCell:=Range("AB80")
'Range("AC120").GoalSeek Goal:=DecOT, ChangingCell:=Range("AC80")


'allowance
'Range("R119").GoalSeek Goal:=50000, ChangingCell:=Range("R104") 'apr
'Range("S119").GoalSeek Goal:=50000, ChangingCell:=Range("S104") 'may
'Range("T119").GoalSeek Goal:=50000, ChangingCell:=Range("T104") 'jun
'Range("U119").GoalSeek Goal:=50000, ChangingCell:=Range("U104")
'Range("V119").GoalSeek Goal:=50000, ChangingCell:=Range("V104")
'Range("W119").GoalSeek Goal:=50000, ChangingCell:=Range("W104")
'Range("X119").GoalSeek Goal:=50000, ChangingCell:=Range("X104")
'Range("Y119").GoalSeek Goal:=50000, ChangingCell:=Range("Y104")
'Range("Z119").GoalSeek Goal:=50000, ChangingCell:=Range("Z104")
'Range("Z119").GoalSeek Goal:=50000, ChangingCell:=Range("Z104")

Range("R83:AC83").ClearContents

Range("R110:AC110").ClearContents

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
    
    If gl = merchGL Then
        tot4 = tot4 * 1.1
        tot5 = tot5 * 1.1
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


