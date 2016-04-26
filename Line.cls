VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All Rights Reserved.

Option Compare Database
Option Explicit

' Private variables for line of text
Private mstrText As String

' Private ID variable
Private mstrID As String

' Public variable for changed flag
Public Changed As Boolean

Property Get Text() As String
    ' Return value of private variable
    Text = mstrText
End Property

Property Let Text(ByVal strText As String)
    ' Change private variable and set changed flag
    mstrText = strText
    Me.Changed = True
End Property

Property Get Length() As Long
    ' Use Len function to return string length
    Length = Len(mstrText)
End Property

Property Get ID() As String
    ' Return value of private variable
    ID = mstrID
End Property

Private Sub Class_Initialize()
    ' Set the object's ID property to a random string
    mstrID = "Line" & CLng(Rnd * (2 ^ 31))
End Sub
