VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' ControlResize Class

' If the READTAGS constant is set to True,
' your project must include the TaggedValues
' and TaggedValue classes, in addition to
' the FormResize, ControlResize, and
' SectionResize classes.

' Setting READTAGS to true allows you to
' control exact behavior of individual controls
' by setting values in their Tag properties.

' Use one of the three following settings:
' FloatIt=<Right|Bottom|Both|None> (default is None)
' SizeIt=<Right|Bottom|Both|None> (default is None)
' ScaleIt=<Yes|True|On|No|False|Off> (default is Yes)

' The settings are generally mutually exclusive, and setting
' more than one to anything besides the default
' value may have unpredictable results. The ScaleIt
' property allows you to override the form's ScaleControls
' property, for individual controls.
#Const READTAGS = True

#If READTAGS Then
' ==================================
' Text constants.
' ==================================
Private Const adhcSizeIt = "SizeIt"
Private Const adhcFloatIt = "FloatIt"
Private Const adhcScaleIt = "ScaleIt"

Private Const adhcSizeRight = "RIGHT"
Private Const adhcSizeBottom = "BOTTOM"
Private Const adhcSizeBoth = "BOTH"
Private Const adhcSizeNone = "NONE"

Private Const adhcFloatRight = "RIGHT"
Private Const adhcFloatBottom = "BOTTOM"
Private Const adhcFloatBoth = "BOTH"
Private Const adhcFloatNone = "NONE"

Private Const adhcScaleYes = "YES"
Private Const adhcScaleTrue = "TRUE"
Private Const adhcScaleOn = "ON"
Private Const adhcScaleNo = "NO"
Private Const adhcScaleFalse = "FALSE"
Private Const adhcScaleOff = "OFF"
Private Const adhcScaleDefault = "DEFAULT"
#End If ' READTAGS

' ==================================
' Enumerated values.
' ==================================
Public Enum ControlSize
    czRight = 1
    czBottom = 2
    czBoth = 3
    czNone = 0
End Enum

Public Enum ControlScale
    csYes = -1
    csNo = 0
    csDefault = 1
End Enum

Public Enum ControlFloat
    cfRight = 1
    cfBottom = 2
    cfBoth = 3
    cfNone = 0
End Enum

' ==================================
' WIN API Declarations
' ==================================

Private Declare Function GetLocaleInfo _
 Lib "kernel32" _
 Alias "GetLocaleInfoA" _
 (ByVal Locale As Long, ByVal LCType As Long, _
 ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Declare Function GetSystemDefaultLCID _
 Lib "kernel32" () As Long

Private Const LOCALE_SLIST = &HC

' ==================================
' Private variables
' ==================================

Private mControl As control
Private alngWidths() As Long
Private mstrColumnWidths As String
Private mfHasColumns As Boolean
Private mstrDelimiter As String
Private mlngOriginalHeight As Long
Private mlngOriginalWidth As Long


' ==================================
' Simple Properties
' ==================================

Public Top As Long
Public Left As Long
Public width As Long
Public height As Long
Public FontSize As Long
Public frm As FormResize
Public Parent As FormResize
Public Name As String

' How should this control be handled?
Public SizeIt As ControlSize
Public FloatIt As ControlFloat
Public ScaleIt As ControlScale

' Offsets from right/bottom.
Public OffsetFromRight As Long
Public OffsetFromBottom As Long

Private Sub Class_Initialize()
    ' Assume reasonable defaults.
    ScaleIt = csDefault
    SizeIt = czNone
    FloatIt = cfNone
End Sub

' ==================================
' Event Procedures
' ==================================

Private Sub Class_Terminate()
    On Error Resume Next
    Call Parent.CleanUp
End Sub

' ==================================
' Public Methods
' (Only public so that other
' classes can call them. In VB,
' these would all be "Friend"
' methods. You should never
' call any of these directly.)
' ==================================

Public Sub CleanUp()
    ' Release references to
    ' other objects.
    Set Parent = Nothing
    Set frm = Nothing
End Sub

Public Property Set ResetControl(ctl As control)
    ' If you lose references (because
    ' you resized the form too small and
    ' Access threw away objects), you need
    ' to be able to reset the pointers.
    ' This procedure does that for you.
    Set mControl = ctl
    If ctl.ControlType = acSubform Then
        Set frm.Parent = Me
        Set frm.ParentControl = ctl
        Set frm.FormReset = ctl.Form
    End If
End Property

Public Property Set control(ctl As control)
    ' Set this object's Control property.
    Set mControl = ctl
    Select Case ctl.ControlType
        Case acListBox, acComboBox
            mfHasColumns = True
        Case acSubform
            Set frm = New FormResize
            Set frm.Parent = Me
            Set frm.ParentControl = ctl
            Set frm.Form = ctl.Form
        Case Else
    End Select
    
#If READTAGS Then
    Dim tv As TaggedValues
    
    If Len(ctl.tag) > 0 Then
        Set tv = New TaggedValues
        tv.Text = ctl.tag
        
        If SizeIt = czNone Then
            Select Case UCase(tv.Item(adhcSizeIt))
                Case adhcSizeRight
                    SizeIt = czRight
                Case adhcSizeBottom
                    SizeIt = czBottom
                Case adhcSizeBoth
                    SizeIt = czBoth
                Case Else
                    SizeIt = czNone
            End Select
        End If
        
        If FloatIt = cfNone Then
            Select Case UCase(tv.Item(adhcFloatIt))
                Case adhcFloatRight
                    FloatIt = cfRight
                Case adhcFloatBottom
                    FloatIt = cfBottom
                Case adhcFloatBoth
                    FloatIt = cfBoth
                Case Else
                    FloatIt = cfNone
            End Select
        End If
        
        If ScaleIt = csDefault Then
            Select Case UCase(tv.Item(adhcScaleIt))
                Case adhcScaleYes, adhcScaleTrue, adhcScaleOn
                    ScaleIt = csYes
                Case adhcScaleNo, adhcScaleFalse, adhcScaleOff
                    ScaleIt = csNo
                Case adhcScaleDefault
                    ScaleIt = csDefault
                Case Else
                    ScaleIt = csDefault
            End Select
        End If
    End If
    Set tv = Nothing
#End If
    
    ' Get offsets in twips.
    mlngOriginalHeight = ctl.height
    mlngOriginalWidth = ctl.width
    OffsetFromBottom = Parent.HeightInTwips - ctl.Top
    OffsetFromRight = Parent.WidthInTwips - ctl.Left
End Property

Public Property Get OriginalHeight() As Long
    OriginalHeight = mlngOriginalHeight
End Property

Public Property Get OriginalWidth() As Long
    OriginalWidth = mlngOriginalWidth
End Property

Public Property Get control() As control
    Set control = mControl
End Property

Public Property Let ColumnWidths(value As String)
    ' Store away this control's column widths,
    ' and parse the delimited string
    ' into an array of values.
    mstrColumnWidths = value
    Call SetColumnWidths(value)
End Property

Public Property Get ColumnWidths() As String
    ' Return the original column widths value.
    ColumnWidths = mstrColumnWidths
End Property

Public Property Get HasColumns() As Boolean
    ' Return back whether this control has columns
    ' (that is, whether it's a ListBox or
    ' ComboBox control) that need to be scaled.
    HasColumns = mfHasColumns
End Property

Public Property Get ScaledColumnWidths(decFactorX As Variant) As String
    Dim i As Integer
    Dim astrNew() As String
    
    ' Take the array of column widths, scale them, and
    ' combine into a string that you can place directly into
    ' the ColumnWidths property.
    
    On Error Resume Next
    ' Join only works with string arrays, so place
    ' scaled column width values into a string array.
    ReDim astrNew(LBound(alngWidths) To UBound(alngWidths))
    If Err.Number <> 0 Then
        ' There's no array here.
        ' Set the ColumnWidths to be the same as
        ' the control width.
        ScaledColumnWidths = control.width
    Else
        For i = LBound(alngWidths) To UBound(alngWidths)
            astrNew(i) = Int(alngWidths(i) * decFactorX)
        Next i
        ScaledColumnWidths = Join(astrNew, mstrDelimiter)
    End If
End Property

' ==================================
' Private methods
' ==================================

Private Sub SetColumnWidths(strValue As String)
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    Dim strItems() As String
    Dim strWidths As String
    Dim i As Integer
    Dim strDelimiter As String
    Dim lngLBound As Long
    Dim lngUBound As Long
    
    ' Stupidly, column widths come as a string. To
    ' work with them individually, you must parse them out,
    ' make changes, and then put them back together.
    ' That's what this procedure does.
        
    On Error Resume Next
    If Len(strValue) > 0 Then
        ' Try to split on semi-colons.
        ' This may fail, if users specified
        ' the widths delimited with the local
        ' delimiter.
        If Len(mstrDelimiter) > 0 Then
            ' If this isn't the first time
            ' you've done this for this control,
            ' then use the delimiter you found on the first
            ' pass through.
            strItems = Split(strValue, mstrDelimiter)
        Else
            strDelimiter = ";"
            strItems = Split(strValue, strDelimiter)
            ' Did the Split return the whole
            ' string as one item? If so,
            ' try splitting on the local delimiter.
            If strItems(0) = strWidths Then
                strDelimiter = GetListSeparator()
                strItems = Split(strWidths, mstrDelimiter)
            End If
            ' Store the delimiter away, so next time
            ' you won't have to go through so much work.
            mstrDelimiter = strDelimiter
        End If
        
        lngLBound = LBound(strItems)
        lngUBound = UBound(strItems)
        ReDim alngWidths(lngLBound To lngUBound)
        For i = lngLBound To lngUBound
            alngWidths(i) = val(strItems(i))
        Next i
    End If
    Err.Clear
End Sub

Private Function GetListSeparator() As String
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    Dim lngLCID As Long
    Dim lngLen As Long
    ' Can't be more than three characters, plus a null terminator.
    Dim strBuffer As String * 4
    
    ' Find the Windows list separator.

    lngLCID = GetSystemDefaultLCID()
    
    ' Get the current ShortDate format.
    lngLen = GetLocaleInfo(lngLCID, LOCALE_SLIST, _
     strBuffer, Len(strBuffer))
    ' Trim off extra junk, and final Null character.
    
    If lngLen > 1 Then
        GetListSeparator = VBA.Left$(strBuffer, lngLen - 1)
    End If
End Function

