Attribute VB_Name = "wageSpread"
Sub GSk()

fourWk = 0.0769
fiveWk = 0.0961

'wage ord
deductAnn = 200188 / 12


JanWO = 63118
FebWO = 76546
MarWO = 89378
AprWO = 77509
MayWO = 77480
JunWO = 98767
JulWO = 77480
AugWO = 77480
SepWO = 98767
OctWO = 77480
NovWO = 77480
DecWO = 98767
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



End Sub


