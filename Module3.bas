Attribute VB_Name = "Module2"
Sub t()

Dim Apr16WO As Long
Dim May16WO As Long
Dim Jun16WO As Long
Dim Jul16WO As Long
Dim Aug16WO As Long
Dim Sep16WO As Long
Dim Oct16WO As Long
Dim Nov16WO As Long
Dim Dec16WO As Long

Apr16WO = 53804.4978751728
May16WO = 53804.4978751728
Jun16WO = 66572.4978751728
Jul16WO = 56492.4978751728
Aug16WO = 47612.4978751728
Sep16WO = 70448.4978751728
Oct16WO = 47612.4978751728
Nov16WO = 43052.4978751728
Dec16WO = 73292.4978751728

Dim AprOT As Long
Dim MayOT As Long
Dim JunOT As Long
Dim JulOT As Long
Dim AugOT As Long
Dim SepOT As Long
Dim OctOT As Long
Dim NovOT As Long
Dim DecOT As Long

AprOT = 5917.64259951416
MayOT = 5274.64259951416
JunOT = 7446.64259951416
JulOT = 3783.64259951416
AugOT = 3715.64259951416
SepOT = 4453.64259951416
OctOT = 2707.64259951416
NovOT = 4333.64259951416
DecOT = 3412.64259951416








Range("R116").GoalSeek Goal:=Apr16WO, ChangingCell:=Range("R81") 'apr
Range("S116").GoalSeek Goal:=May16WO, ChangingCell:=Range("S81") 'may
Range("T116").GoalSeek Goal:=Jun16WO, ChangingCell:=Range("T81") 'jun
Range("U116").GoalSeek Goal:=Jul16WO, ChangingCell:=Range("U81")
Range("V116").GoalSeek Goal:=Aug16WO, ChangingCell:=Range("V81")
Range("W116").GoalSeek Goal:=Sep16WO, ChangingCell:=Range("W81")
Range("X116").GoalSeek Goal:=Oct16WO, ChangingCell:=Range("X81")
Range("Y116").GoalSeek Goal:=Nov16WO, ChangingCell:=Range("Y81")
Range("Z116").GoalSeek Goal:=50000, ChangingCell:=Range("Z81")
Range("Z116").GoalSeek Goal:=Dec16WO, ChangingCell:=Range("Z81")

'o/t
Range("R120").GoalSeek Goal:=AprOT, ChangingCell:=Range("R80") 'apr
Range("S120").GoalSeek Goal:=MayOT, ChangingCell:=Range("S80") 'may
Range("T120").GoalSeek Goal:=JunOT, ChangingCell:=Range("T80") 'jun
Range("U120").GoalSeek Goal:=JulOT, ChangingCell:=Range("U80")
Range("V120").GoalSeek Goal:=AugOT, ChangingCell:=Range("V80")
Range("W120").GoalSeek Goal:=SepOT, ChangingCell:=Range("W80")
Range("X120").GoalSeek Goal:=OctOT, ChangingCell:=Range("X80")
Range("Y120").GoalSeek Goal:=NovOT, ChangingCell:=Range("Y80")
Range("Z120").GoalSeek Goal:=DecOT, ChangingCell:=Range("Z80")
Range("Z120").GoalSeek Goal:=50000, ChangingCell:=Range("Z80")



End Sub
