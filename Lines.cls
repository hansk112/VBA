VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All Rights Reserved.

Option Compare Database
Option Explicit

' Private collection to store Lines
Private mcolLines As Collection

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    ' Enumeration function for the collection
    Set NewEnum = mcolLines.[_NewEnum]
End Function

Private Sub Class_Initialize()
    ' Instantiate the Collection object
    Set mcolLines = New Collection
End Sub

Public Sub Add(ByVal strText As String, Optional ByVal varBefore As Variant)
    ' Declare new Line object
    Dim objLine As New Line
    
    ' Set Text property to passed string
    objLine.Text = strText
    ' Add to private collection, using object's
    ' ID property as unique index
    mcolLines.Add objLine, objLine.ID, varBefore
End Sub

Public Sub Remove(ByVal varID As Variant)
    ' Call Remove method of private collection object
    mcolLines.Remove varID
End Sub

Public Function Item(ByVal varID As Variant) As Line
Attribute Item.VB_UserMemId = 0
    ' Set return value of property to item within
    ' the private collection object specified by
    ' the passed index value (Note the return type!)
    Set Item = mcolLines(varID)
End Function

Property Get Count() As Long
    ' Return Count property of private collection
    Count = mcolLines.Count
End Property

Property Let Changed(ByVal fChanged As Boolean)
    Dim objLine As Line
    
    ' Set Changed property of each Line to fChanged
    For Each objLine In mcolLines
        objLine.Changed = fChanged
    Next
End Property

Property Get Changed() As Boolean
    Dim objLine As Line

    ' Loop through all Line objects in collection--
    ' if any Changed property is True then the
    ' Changed property of the collection is True
    For Each objLine In mcolLines
        If objLine.Changed Then
            Changed = True
            Exit For
        End If
    Next
End Property
