Attribute VB_Name = "basCommon"
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All rights reserved.

' Required by:
'   basFileOpen
'   basFontHandling
'   basObjList
'   CommonDlg
'   ScreenInfo
'   ShellBrowse
'   VersionInfo

' Common routines needed by many of the procedures
' in this project.

' Registry Errors, used in several modules.
Public Enum adhRegErrors
    adhcAccErrSuccess = 0
    adhcAccErrUnknown = -1
    adhcAccErrRegKeyNotFound = -201
    adhcAccErrRegValueNotFound = -202
    adhcAccErrRegCantSetValue = -203
    adhcAccErrRegSubKeyNotFound = -204
    adhcAccErrRegTypeNotSupported = -205
    adhcAccErrRegCantCreateKey = -206
    adhcAccErrRegBufferTooSmall = -207
    adhcAccErrRegCantDeleteValue = -208
End Enum

' Indicate that a parameter for QuickSort is missing.
Private Const dhcMissing = -2

Public Function adhFnPtrToLong(lngAddress As Long) As Long
    
    ' Given a function pointer as a Long, return a Long.
    ' Sure looks like this function isn't doing anything,
    ' and in reality, it's not.
    
    ' Call this function like this:
    '
    ' lngPointer = adhFnPtrToLong(AddressOf SomeFunction)
    
    ' and it returns the address you've sent it
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All rights reserved.

    ' In:
    '   lngAddress:
    '       address of a public procedure, passed using
    '       the AddressOf modifier.
    ' Out:
    '   Return value:
    '       The input address, cast as a Long.

    adhFnPtrToLong = lngAddress
End Function

Public Function adhTrimNull(strVal As String) As String
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All rights reserved.
    
    ' Trim the end of a string, stopping at the first
    ' null character.
    
    Dim intPos As Integer
    intPos = InStr(1, strVal, vbNullChar)
    Select Case intPos
        Case Is > 1
            adhTrimNull = Left$(strVal, intPos - 1)
        Case 0
            adhTrimNull = strVal
        Case 1
            adhTrimNull = vbNullString
    End Select
End Function

Public Sub adhQuickSort(varArray As Variant, _
 Optional intLeft As Integer = dhcMissing, _
 Optional intRight As Integer = dhcMissing)

    ' Quicksort for simple data types.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All rights reserved.
    
    ' Originally from "VBA Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 1997; Sybex, Inc. All rights reserved.
    
    ' Entry point for sorting the array.
    
    ' This technique uses the recursive Quicksort
    ' algorithm to perform its sort.
    
    ' In:
    '   varArray:
    '       A variant pointing to an array to be sorted.
    '       This had better actually be an array, or the
    '       code will fail, miserably. You could add
    '       a test for this:
    '       If Not IsArray(varArray) Then Exit Sub
    '       but hey, that would slow this down, and it's
    '       only YOU calling this procedure.
    '       Make sure it's an array. It's your problem.
    '   intLeft:
    '   intRight:
    '       Lower and upper bounds of the array to be sorted.
    '       If you don't supply these values (and normally, you won't)
    '       the code uses the LBound and UBound functions
    '       to get the information. In recursive calls
    '       to the sort, the caller will pass this information in.
    '       To allow for passing integers around (instead of
    '       larger, slower variants), the code uses -2 to indicate
    '       that you've not passed a value. This means that you won't
    '       be able to use this mechanism to sort arrays with negative
    '       indexes, unless you modify this code.
    ' Out:
    '       The data in varArray will be sorted.
    
    Dim i As Integer
    Dim j As Integer
    Dim varTestVal As Variant
    Dim intMid As Integer

    If intLeft = dhcMissing Then intLeft = LBound(varArray)
    If intRight = dhcMissing Then intRight = UBound(varArray)
   
    If intLeft < intRight Then
        intMid = (intLeft + intRight) \ 2
        varTestVal = varArray(intMid)
        i = intLeft
        j = intRight
        Do
            Do While varArray(i) < varTestVal
                i = i + 1
            Loop
            Do While varArray(j) > varTestVal
                j = j - 1
            Loop
            If i <= j Then
                Call SwapElements(varArray, i, j)
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= intMid Then
            Call adhQuickSort(varArray, intLeft, j)
            Call adhQuickSort(varArray, i, intRight)
        Else
            Call adhQuickSort(varArray, i, intRight)
            Call adhQuickSort(varArray, intLeft, j)
        End If
    End If
End Sub

Private Sub SwapElements(varItems As Variant, intItem1 As Integer, intItem2 As Integer)
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All rights reserved.
    
    Dim varTemp As Variant

    varTemp = varItems(intItem2)
    varItems(intItem2) = varItems(intItem1)
    varItems(intItem1) = varTemp
End Sub
