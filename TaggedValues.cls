VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaggedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Compare Database
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' This class implements a collection
' of tag/value pairs.

Private mstrSeparator As String
Private mcolItems As Collection

Private Sub Class_Initialize()
    ' Assume ";" as the separator.
    mstrSeparator = ";"
    Set mcolItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolItems = Nothing
End Sub

Public Property Get Count() As Long
    Count = mcolItems.Count
End Property

Private Sub ParseItems(strText As String)
    ' Look for mstrSeparator, and break up the string into
    ' x=y pairs. Then parse each of those, looking
    ' for equal signs.
    Dim varItems As Variant
    Dim varText As Variant
    Dim i As Integer
    Dim strTag As String
    Dim strValue As String
    
    varItems = Split(strText, mstrSeparator)
    For i = LBound(varItems) To UBound(varItems)
        varText = Split(varItems(i), "=")
        strTag = varText(LBound(varText))
        strValue = varText(UBound(varText))
        Call Add(strTag, strValue)
    Next i
End Sub

Public Property Get Separator() As String
    Separator = mstrSeparator
End Property

Public Property Let Separator(value As String)
    mstrSeparator = value
End Property

Public Property Get Exists(tag As String) As Boolean
    ' Does this particular tag already exist in
    ' this collection of tags?
    On Error Resume Next
    Dim strValue As String
    
    ' If you can retrieve the value, the tag already
    ' exists.
    strValue = mcolItems.Item(tag).value
    Exists = (Err.Number = 0)
    Err.Clear
End Property

Public Function Item(tag As String) As String
    If Exists(tag) Then
        Item = mcolItems.Item(tag).value
    End If
End Function

Public Sub Remove(tag As String)
    ' Check first to make sure the tag exists.
    ' If so, remove it.
    If Exists(tag) Then
        mcolItems.Remove tag
    End If
End Sub

Public Function Add(tag As String, _
 value As String) As TaggedValue
    Dim tv As TaggedValue
    
    ' If the tag already exists, remove
    ' it so you can set the new value.
    Call Remove(tag)
    
    ' Create the new TaggedValue object, set its
    ' properties, and add it to the collection.
    Set tv = New TaggedValue
    tv.tag = tag
    tv.value = value
    mcolItems.Add tv, tag
End Function

Public Property Let Text(value As String)
    ' Parse all the values out into
    ' text/value pairs.
    
    ' Look for mstrSeparator, and break up
    '  the string into x=y pairs. Then parse
    ' each of those, looking for equal signs.
    Dim varItems As Variant
    Dim varText As Variant
    Dim i As Integer
    Dim strTag As String
    Dim strValue As String
    
    varItems = Split(value, mstrSeparator)
    For i = LBound(varItems) To UBound(varItems)
        varText = Split(varItems(i), "=")
        strTag = varText(LBound(varText))
        strValue = varText(UBound(varText))
        Call Add(strTag, strValue)
    Next i
End Property

Public Property Get Text() As String
    Dim tv As TaggedValue
    Dim strOut As String
    
    ' Loop through each item in the collection.
    ' For any that have non-empty Tag properties
    ' (and they all really should), tack on the
    ' Tag=Value pairs.
    For Each tv In mcolItems
        If Len(tv.tag) > 0 Then
            strOut = strOut & mstrSeparator & _
             tv.tag & "=" & tv.value
        End If
    Next tv
    ' If there's anything in the output string,
    ' it'll have a leading separator.  Remove that now.
    If Len(strOut) > Len(mstrSeparator) Then
        strOut = Mid$(strOut, Len(mstrSeparator) + 1)
    End If
    Text = strOut
End Property

