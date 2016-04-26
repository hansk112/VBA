Attribute VB_Name = "bpcSpreadVar"
Sub sprVar()


Dim apr As Long, may As Long, jun As Long, jul As Long, aug As Long, sep As Long, oct As Long, nov As Long, dec As Long
Dim rng As Range
Dim actrow As Long
Dim actcol As Long

Set rng = Selection
actrow = ActiveCell.Row
actcol = ActiveCell.Column
var = ActiveCell
Debug.Print var



apr = Range("V" & actrow)
may = Range("W" & actrow)
jun = Range("X" & actrow)
jul = Range("Y" & actrow)
aug = Range("Z" & actrow)
sep = Range("AA" & actrow)
oct = Range("AB" & actrow)
nov = Range("AC" & actrow)
dec = Range("AD" & actrow)
        
If var < 0 Then

var = var / 9
apr = var + apr
may = var + may
jun = var + jun
jul = var + jul
aug = var + aug
sep = var + sep
oct = var + oct
nov = var + nov
dec = var + dec

End If
                
If var < 0 Then

Range("V" & actrow) = apr
Range("W" & actrow) = may
Range("X" & actrow) = jun
Range("Y" & actrow) = jul
Range("Z" & actrow) = aug
Range("AA" & actrow) = sep
Range("AB" & actrow) = oct
Range("AC" & actrow) = nov
Range("AD" & actrow) = dec

End If



End Sub

Sub tst()

Dim rng As Range

Set rng = Selection
Debug.Print rng

End Sub

Sub SelectActualUsedRange()
  Dim FirstCell As Range, LastCell As Range
  Set LastCell = Cells(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column)
  Set FirstCell = Cells(Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlRows, _
      SearchDirection:=xlNext, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlByColumns, _
      SearchDirection:=xlNext, LookIn:=xlValues).Column)
  Range(FirstCell, LastCell).Select
  Debug.Print FirstCell & LastCell
End Sub

Sub rGrp()




End Sub
