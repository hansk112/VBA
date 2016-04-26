VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' SectionResize Class

' Keep track of form Section information.
' This class only exists because you can't
' put a user-defined type into a collection.

Public height As Long
Public Section As Integer
