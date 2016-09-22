Attribute VB_Name = "Module1"
Sub updateTable()

Dim nrows As Long, nrowsGrocery As Long

Sheet5.Activate

nrows = Cells(Rows.Count, 7).End(xlUp).Row

Range("G16:M" & nrows).Copy


Sheet2.Activate

nrowsGrocery = Cells(Rows.Count, 7).End(xlUp).Row
nrowsAddOne = nrowsGrocery + 1

Range("G" & nrowsAddOne).PasteSpecial xlPasteValues

Worksheets("Table").Activate

Call chngDate


End Sub
Sub chngDate()

Dim nrowsPost As Long

p1 = "p1 Sep-15"
p2 = "p2 Oct-15"
p3 = "p3 Nov-15"
p4 = "p4 Dec-15"
p5 = "p5 Jan-16"
p6 = "p6 Feb-16"
p7 = "p7 Mar-16"
p8 = "p8 Apr-16"
p9 = "p9 May-16"
p10 = "p10 Jun-16"
p11 = "p11 Jul-16"
p12 = "p12 Aug-16"
Dim rng As Range

Worksheets("Table").Activate
nrowsPost = Cells(Rows.Count, 10).End(xlUp).Row
Set rng = Range("J15:J" & nrowsPost)

' sht.Cells.Replace what:=fnd, Replacement:=rplc, _

With ActiveSheet
    rng.Cells.Replace What:="Sep", Replacement:=p1
    rng.Cells.Replace What:="Oct", Replacement:=p2
    rng.Cells.Replace What:="Nov", Replacement:=p3
    rng.Cells.Replace What:="Dec", Replacement:=p4
    rng.Cells.Replace What:="Jan", Replacement:=p5
    rng.Cells.Replace What:="Feb", Replacement:=p6
    rng.Cells.Replace What:="Mar", Replacement:=p7
    rng.Cells.Replace What:="Apr", Replacement:=p8
    rng.Cells.Replace What:="May", Replacement:=p9
    rng.Cells.Replace What:="Jun", Replacement:=p10
    rng.Cells.Replace What:="Jul", Replacement:=p11
    rng.Cells.Replace What:="Aug", Replacement:=p12
End With
    
    
    
    
End Sub

