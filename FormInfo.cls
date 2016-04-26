VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' FormInfo Class module

' Generally, leave this False. If you want to
' debug odd behaviors, change it to True.
#Const DEBUGGING = False

' ==================================
' Windows API declarations.
' ==================================

Private Declare Function SendMessage Lib "user32" _
 Alias "SendMessageA" _
 (ByVal hwnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
 Alias "SetWindowLongA" (ByVal hwnd As Long, _
 ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
 Alias "GetWindowLongA" (ByVal hwnd As Long, _
 ByVal nIndex As Long) As Long

Private Declare Function ShowWindow _
 Lib "user32" _
 (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function MoveWindow _
 Lib "user32" _
 (ByVal hwnd As Long, _
 ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect _
 Lib "user32" _
 (ByVal hwnd As Long, _
 lpRect As RECT) As Long

Private Declare Function GetParent _
 Lib "user32" _
 (ByVal hwnd As Long) As Long

Private Declare Function IsIconic _
 Lib "user32" _
 (ByVal hwnd As Long) As Long
 
 Private Declare Function IsZoomed _
  Lib "user32" _
  (ByVal hwnd As Long) As Long

Private Declare Function GetClientRect _
 Lib "user32" _
 (ByVal hwnd As Long, _
 lpRect As RECT) As Long

Private Declare Function GetDeviceCaps _
 Lib "gdi32" _
 (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Declare Function GetSystemMetrics _
 Lib "user32" _
 (ByVal nIndex As Long) As Long
 
Private Declare Function GetDC _
 Lib "user32" _
 (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC _
 Lib "user32" _
 (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Declare Function SystemParametersInfoRect _
 Lib "user32" Alias "SystemParametersInfoA" _
 (ByVal uAction As Long, ByVal uParam As Long, _
 ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
 
Private Declare Function ClientToScreen _
 Lib "user32" _
 (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const SPI_GETWORKAREA = 48

Private Const SM_CYCAPTION = 4
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1

' GetWindowLong Constant
Private Const GWL_STYLE = -16
Private Const WS_SYSMENU = &H80000

' Windows Style constant
Private Const WS_CAPTION = &HC00000

' Windows message constant.
Private Const WM_NCPAINT = &H85

' Store rectangle coordinates.
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' ==================================
' Constants
' ==================================

' Windows 95/98/NT4/2000 puts a 2-pixel
' border around the MDI client area, which
' doesn't get taken into account automatically.
' If you're using NT 3.51, you're on your own.
Private Const adhcBorderWidthX = 2
Private Const adhcBorderWidthY = 2

Private Const adhcTop = "Top"
Private Const adhcLeft = "Left"
Private Const adhcRight = "Right"
Private Const adhcBottom = "Bottom"
Private Const adhcWidth = "Width"
Private Const adhcHeight = "Height"
Private Const adhcState = "State"

Private Const adhcTwipsPerInch = 1440

' ==================================
' Enums
' ==================================

Private Enum ReferenceType
    rtPopup = 0
    rtNormal = 1
End Enum

Public Enum WindowState
    wsNormal = 0
    wsMinimized = 1
    wsMaximized = 2
End Enum

' ==================================
' Private variables
' ==================================

Private frm As Form

Private mptCurrentScreen As POINTAPI
Private mptTwipsPerPixel As POINTAPI
Private mptClientOffset As POINTAPI

' Store form coordinates.
Private Type COORDS
    Left As Long
    Top As Long
    width As Long
    height As Long
    State As WindowState
End Type

Private mCoords As COORDS
Private mfIsSubform As Boolean
Private mfIsPopup As Boolean


' ==================================
' Public Methods and Properties
' ==================================

Public Property Get ScreenX() As Long
    ' Return the current screen size, in pixels.
    ScreenX = mptCurrentScreen.x
End Property

Public Property Get ScreenY() As Long
    ' Return the current screen size, in pixels.
    ScreenY = mptCurrentScreen.y
End Property

Public Property Get ScreenXinTwips() As Long
    ' Return the current screen size, in pixels.
    ScreenXinTwips = mptCurrentScreen.x * mptTwipsPerPixel.x
End Property

Public Property Get ScreenYInTwips() As Long
    ' Return the current screen size, in pixels.
    ScreenYInTwips = mptCurrentScreen.y * mptTwipsPerPixel.y
End Property

Public Property Get WindowState() As WindowState
    ' Retrieve the current window state
    ' as a single value.
    On Error GoTo HandleErrors
    If IsMaximized Then
        WindowState = wsMaximized
    ElseIf IsMinimized Then
        WindowState = wsMinimized
    Else
        WindowState = wsNormal
    End If

ExitHere:
    Exit Property
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WindowState", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let WindowState(value As WindowState)
    On Error GoTo HandleErrors
    Select Case value
        Case wsMaximized
            IsMaximized = True
        Case wsMinimized
            IsMinimized = True
        Case wsNormal
            ' Just set one of minimized
            ' or maximized to False. That
            ' will restore the window.
            IsMaximized = False
    End Select
    
ExitHere:
    Exit Property
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WindowState", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Sub RetrieveCoords(AppName As String)

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    Dim strName As String
    Dim C As COORDS

    On Error GoTo HandleErrors

    ' Use the name of the application as the highest
    ' level, and the form's name as the next level.
    ' This way, you could have multiple forms in the same
    ' app use this code.
    strName = frm.Name
    With C
        .State = val(GetSetting(AppName, _
         strName, adhcState, Me.WindowState))
        Select Case .State
            Case wsMinimized
                IsMinimized = True
            Case wsNormal
                .Top = _
                 GetSetting(AppName, strName, adhcTop, .Top)
                .Left = _
                 GetSetting(AppName, strName, adhcLeft, .Left)
                .width = _
                 GetSetting(AppName, strName, adhcWidth, .width)
                .height = _
                 GetSetting(AppName, strName, adhcHeight, .height)
        
                ' Only muck with the form's size if
                ' you get values for width and height
                ' that make sense.
                If .width > 0 And .height > 0 Then
                    Call SetSize(.Left, .Top, .width, .height)
                End If
            Case wsMaximized
                ' Don't set this form to be maximized
                ' or not. If the MDI Client window is maximized
                ' you want to bypass these settings and just
                ' maximize the window.
        End Select
    End With
   
ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.RetrieveCoords", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Sub SaveCoords(AppName As String)

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    Dim strName As String

    On Error GoTo HandleErrors

    strName = frm.Name
    
    ' Use the name of the application as the highest
    ' level, and the form's name as the next level.
    ' This way, you could have multiple forms in the same
    ' app use this code.
    
    Call SaveSetting( _
     AppName, strName, adhcState, WindowState)
    Call SaveSetting( _
     AppName, strName, adhcTop, Top)
    Call SaveSetting( _
     AppName, strName, adhcLeft, Left)
    Call SaveSetting( _
     AppName, strName, adhcWidth, width)
    Call SaveSetting( _
     AppName, strName, adhcHeight, height)
    
ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.SaveCoords", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Property Let PercentWidth(value As Double)
    ' Size the window to the specified percentage
    ' of the available screen width.
    On Error GoTo HandleErrors
    Dim cParent As COORDS
    
    Call GetCoords(frm.hwnd, mCoords)
    Call GetParentCoords(frm.hwnd, cParent)
    mCoords.width = cParent.width * value / 100
    Call SetSize(width:=mCoords.width)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentWidth", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get PercentWidth() As Double
    On Error GoTo HandleErrors
    
    Dim cParent As COORDS
    Call GetCoords(frm.hwnd, mCoords)
    Call GetParentCoords(frm.hwnd, cParent)
    PercentWidth = mCoords.width / cParent.width * 100

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentWidth", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let PercentHeight(value As Double)
    ' Size the window to the specified percentage
    ' of the available screen Height.
    On Error GoTo HandleErrors
    Dim cParent As COORDS
    
    Call GetCoords(frm.hwnd, mCoords)
    Call GetParentCoords(frm.hwnd, cParent)
    mCoords.height = cParent.height * value / 100
    Call SetSize(height:=mCoords.height)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentHeight", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get PercentHeight() As Double
    On Error GoTo HandleErrors
    
    Dim cParent As COORDS
    Call GetCoords(frm.hwnd, mCoords)
    Call GetParentCoords(frm.hwnd, cParent)
    PercentHeight = mCoords.height / cParent.height * 100

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentHeight", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get IsPopup() As Boolean
    ' You can't tell from properties if a
    ' form is a Popup form or not -- you
    ' can either set the popup property to
    ' True, or you can open it with the
    ' acDialog flag. This property
    ' checks the parent of the form, and
    ' if it's Access, this is a popup form.
    IsPopup = mfIsPopup
End Property


Public Property Get IsSubForm() As Boolean
    IsSubForm = mfIsSubform
End Property

Public Property Get IsMaximized() As Boolean
    On Error GoTo HandleErrors
    IsMaximized = IsZoomed(frm.hwnd)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMaximized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let IsMaximized(MaxIt As Boolean)
    On Error GoTo HandleErrors
    If MaxIt Then
        Call ShowWindow(frm.hwnd, SW_MAXIMIZE)
    Else
        Call ShowWindow(frm.hwnd, SW_NORMAL)
    End If

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMaximized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get IsMinimized() As Boolean
    On Error GoTo HandleErrors
    IsMinimized = IsIconic(frm.hwnd)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMinimized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let IsMinimized(MinIt As Boolean)
    On Error GoTo HandleErrors
    If MinIt Then
        Call ShowWindow(frm.hwnd, SW_MINIMIZE)
    Else
        Call ShowWindow(frm.hwnd, SW_NORMAL)
    End If

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMinimized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Set Form(value As Form)
    On Error GoTo HandleErrors
    Set frm = value
    Call GetScreenInfo
    Call GetClientOffsets
    mfIsPopup = GetIsPopup()
    mfIsSubform = GetIsSubform()
    
ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Form", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get Form() As Form
    On Error GoTo HandleErrors
    Set Form = frm

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Form", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get TwipsPerPixelX() As Long
    ' Return the screen's vertical Twips/Pixel ratio.
    On Error GoTo HandleErrors
    TwipsPerPixelX = mptTwipsPerPixel.x

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TwipsPerPixelX", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get TwipsPerPixelY() As Long
    ' Return the screen's horizontal Twips/Pixel ratio.
    On Error GoTo HandleErrors
    TwipsPerPixelY = mptTwipsPerPixel.y

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TwipsPerPixelY", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Sub GetSize( _
 Optional ByRef Left As Long, Optional ByRef Top As Long, _
 Optional ByRef width As Long, Optional ByRef height As Long, _
 Optional ByVal InTwips As Boolean = False)
   
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ' Get all the coords at once.
    ' If you want more than one coordinate,
    ' this will be slightly more efficient than
    ' requesting each one individually.
    On Error GoTo HandleErrors
    Dim intMultiplierX As Integer
    Dim intMultiplierY As Integer
    
    If InTwips Then
        intMultiplierX = mptTwipsPerPixel.x
        intMultiplierY = mptTwipsPerPixel.y
    Else
        intMultiplierX = 1
        intMultiplierY = 1
    End If
    Call GetCoords(frm.hwnd, mCoords)
    
    With mCoords
        Left = .Left * intMultiplierX
        Top = .Top * intMultiplierY
        width = .width * intMultiplierX
        height = .height * intMultiplierY
    End With

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetSize", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Sub SetSize( _
 Optional ByVal Left As Variant, _
 Optional ByVal Top As Variant, _
 Optional ByVal width As Variant, _
 Optional ByVal height As Variant, _
 Optional ByVal InTwips As Boolean = False)

    ' Set the form's location/size, either in
    ' pixels (the default) or twips. If you specify
    ' twips, it's the same as using DoCmd.MoveSize,
    ' except that this can work with any form, not just
    ' the current form. MoveSize is limited to
    ' the current form, only.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        If IsMissing(Left) Then
            Left = .Left
        End If
        If Not IsNumeric(Left) Then
            Left = .Left
        End If
        If InTwips Then
            Left = Left \ mptTwipsPerPixel.x
        End If
        
        If IsMissing(Top) Then
            Top = .Top
        End If
        If Not IsNumeric(Top) Then
            Top = .Top
        End If
        If InTwips Then
            Top = Top \ mptTwipsPerPixel.y
        End If
        
        If IsMissing(width) Then
            width = .width
        End If
        If Not IsNumeric(width) Then
            width = .width
        End If
        If InTwips Then
            width = width \ mptTwipsPerPixel.x
        End If
        
        If IsMissing(height) Then
            height = .height
        End If
        If Not IsNumeric(height) Then
            height = .height
        End If
        If InTwips Then
            height = height \ mptTwipsPerPixel.y
        End If
    End With
    Call MoveWindow(frm.hwnd, _
     Left, Top, width, height, 1)

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.SetSize", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Property Get Left() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    Left = mCoords.Left

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Left", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let Left(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         value, .Top, .width, .height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Left", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get LeftInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    LeftInTwips = mCoords.Left * mptTwipsPerPixel.x

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.LeftInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let LeftInTwips(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         value \ mptTwipsPerPixel.x, .Top, .width, .height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.LeftInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get Top() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    Top = mCoords.Top

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Top", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let Top(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         .Left, value, .width, .height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Top", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get TopInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    TopInTwips = mCoords.Top * mptTwipsPerPixel.y

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TopInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let TopInTwips(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         .Left, value \ mptTwipsPerPixel.y, .width, .height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TopInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get width() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    width = mCoords.width

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Width", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let width(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         .Left, .Top, value, .height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Width", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get WidthInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    WidthInTwips = mCoords.width * mptTwipsPerPixel.x

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WidthInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let WidthInTwips(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         .Left, .Top, value \ mptTwipsPerPixel.x, .height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WidthInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property


Public Property Get height() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    height = mCoords.height

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Height", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let height(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         .Left, .Top, .width, value, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Height", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get HeightInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    HeightInTwips = mCoords.height * mptTwipsPerPixel.y

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.HeightInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let HeightInTwips(value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hwnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hwnd, _
         .Left, .Top, .width, value \ mptTwipsPerPixel.y, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.HeightInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Sub FillClientArea()
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    ' Move a form so that it fills the entire
    ' MDI Client area in Access. This allows you to
    ' remove the Close button (or the entire
    ' caption bar, if you like) and "maximize" the
    ' form without Access adding those pesky buttons
    ' back.

    Dim hWndParent As Long
    Dim cParent As COORDS
    
    On Error GoTo HandleErrors
        
    If IsZoomed(frm.hwnd) Then
        Call ShowWindow(frm.hwnd, SW_NORMAL)
    End If
    Call GetParentCoords(frm.hwnd, cParent)
    With cParent
        Call MoveWindow(frm.hwnd, _
         .Left, .Top, .width, .height, 1)
    End With
    
ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.FillClientArea", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere

End Sub

Public Sub Center()

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ' Center a form within the confines of the
    ' MDI parent window. If the form is a popup
    ' form, then center on the screen.
    
    On Error GoTo HandleErrors
    Dim cParent As COORDS

    ' Get the coordinates of the current form.
    Call GetCoords(frm.hwnd, mCoords)
    Call GetParentCoords(frm.hwnd, cParent)
    
    ' Calculate the width of the child form,
    ' and calculate its new coordinates, relative
    ' to its parent.
    With mCoords
        .Left = (cParent.width - .width) \ 2
        .Top = (cParent.height - .height) \ 2
        
        ' Move the child window to its new location.
        Call MoveWindow(frm.hwnd, _
         .Left, .Top, .width, .height, 1)
    End With

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Center", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Property Let ShowSystemMenu(ShowIt As Boolean)
    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
    Dim lngNewStyle As Long
    
    ' Show or hide a form's system menu, depending
    ' on the value in ShowIt.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ' If there is nothing to do, get out.
    If Me.ShowSystemMenu = ShowIt Then
        Exit Property
    End If

    ' Get the current window style of the form.
    lngOldStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    
    If ShowIt Then
    ' Turn on the bit that enables system menu.
    lngNewStyle = lngOldStyle Or WS_SYSMENU
    Else
    ' Turn off the bit the shows the system menu.
    lngNewStyle = lngOldStyle And Not WS_SYSMENU
    End If
    
    ' Set the new window style.
    Call SetWindowLong(frm.hwnd, GWL_STYLE, lngNewStyle)
    
    ' The 1 as the third parameter tells
    ' the window to repaint its entire border.
    Call SendMessage(frm.hwnd, WM_NCPAINT, 1, 0)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowSystemMenu", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get ShowSystemMenu() As Boolean
    
    ' Retrieve info about a form's system menu.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
  
    lngOldStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    ShowSystemMenu = ((lngOldStyle And WS_SYSMENU) = WS_SYSMENU)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowSystemMenu", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let ShowCaptionBar(ShowIt As Boolean)

    ' Show or remove a form's caption bar, depending
    ' on the value in ShowIt.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
    Dim lngNewStyle As Long
    Dim rct As RECT
    Dim intDiff As Integer
    
    ' If there is nothing to do, get out.
    If Me.ShowCaptionBar = ShowIt Then
        GoTo ExitHere
    End If

    Call GetCoords(frm.hwnd, mCoords)
    
    ' Get the current window style of the form.
    lngOldStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    
    If ShowIt Then
        ' Turn off the bit that enables the caption.
        lngNewStyle = lngOldStyle Or WS_CAPTION
    Else
        ' Turn off the bit that enables the caption.
        lngNewStyle = lngOldStyle And Not WS_CAPTION
    End If
    
    ' Set the new window style.
    lngOldStyle = SetWindowLong(frm.hwnd, _
     GWL_STYLE, lngNewStyle)
    
    ' How much room does that caption take up?
    intDiff = GetSystemMetrics(SM_CYCAPTION)
    
    ' Calculate the new height.
    If ShowIt Then
        mCoords.height = mCoords.height + intDiff
    Else
        mCoords.height = mCoords.height - intDiff
    End If

    ' Move the window to the same left and top,
    ' but with new width and height.
    ' This will make the new form appear
    ' a little shorter or a little taller.
    With mCoords
        Call MoveWindow(frm.hwnd, _
         .Left, .Top, .width, .height, 1)
    End With
    Call GetClientOffsets
    
ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowCaptionBar", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get ShowCaptionBar() As Boolean

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
    Dim lngNewStyle As Long
    Dim rct As RECT
    Dim intDX As Integer
    Dim intDY As Integer

    ' Get the current window style of the form.
    lngOldStyle = GetWindowLong(frm.hwnd, GWL_STYLE)

    ' Turn off the bit that enables the caption.
    ShowCaptionBar = ((lngOldStyle And WS_CAPTION) = WS_CAPTION)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowCaptionBar", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get ClientOffsetXinTwips() As Long
    ' Retrieve the client offset, converted to twips.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetXinTwips = mptClientOffset.x * mptTwipsPerPixel.x
End Property

Public Property Get ClientOffsetYinTwips() As Long
    ' Retrieve the client offset, converted to twips.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetYinTwips = mptClientOffset.y * mptTwipsPerPixel.y
End Property

Public Property Get ClientOffsetX() As Long
    ' Retrieve the client offset, in pixels.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetX = mptClientOffset.x
End Property

Public Property Get ClientOffsetY() As Long
    ' Retrieve the client offset, in pixels.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetY = mptClientOffset.y
End Property

' ==================================
' Private Methods
' ==================================

Private Sub GetCoords(hwnd As Long, cToFill As COORDS)

    ' Fill in rct with the coordinates of the client area.
    ' Fill in rctWindow with coordinates of the window.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors
    Dim hWndParent As Long
    Dim rctParent As RECT
    Dim rct As RECT

    ' Find the position of the window in question, in
    ' relation to its parent window (the Access desktop,
    ' the MDIClient window).
    hWndParent = GetParent(hwnd)

   ' Get the coordinates of the current window and its parent.
    Call GetWindowRect(hwnd, rct)

    ' Catch the case where the form is Popup (that is,
    ' its parent is NOT the Access main window.)  In that
    ' case, don't subtract off the coordinates of the
    ' Access MDIClient window.
    Select Case ReferenceType
        Case rtPopup
            ' No special calculations necessary.
        Case rtNormal
            ' Get the MDI Client window's coordinates.
            Call GetWindowRect(hWndParent, rctParent)
    
            ' Subtract off the left and top parent coordinates,
            ' since you need coordinates relative to the parent
            ' for the MoveWindow function call.
            ' Also, subtract off the small border of the
            ' MDI Client window that no one will admit to.
            With rct
                .Left = .Left - rctParent.Left - adhcBorderWidthX
                .Top = .Top - rctParent.Top - adhcBorderWidthY
                .Right = .Right - rctParent.Left - adhcBorderWidthX
                .Bottom = .Bottom - rctParent.Top - adhcBorderWidthY
            End With
    End Select
    cToFill = RctToCoords(rct)

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetCoords", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Sub HandleError(strName As String, _
 lngNumber As Long, strDescription As String)
    
#If DEBUGGING Then
    MsgBox "Error: " & strDescription & _
     " (" & lngNumber & ")", vbExclamation, strName
    ' Trigger a breakpoint. Remove this
    ' if you don't want a breakpoint here.
    Debug.Assert False
#End If
End Sub

Private Sub GetScreenInfo()
    ' This procedure fills in the module variables
    ' mptCurrentScreen, mptTwipsPerPixel
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors
    Dim lngDC As Long
    Dim ptCurrentDPI As POINTAPI
    Const HWND_DESKTOP = 0
    
    lngDC = GetDC(HWND_DESKTOP)
    
    ' If the call to GetDC didn't fail (and it had
    ' better not, or things are really busted),
    ' then get the info.
    
    If lngDC <> 0 Then
        ' Find the number of pixels in both directions
        ' on the screen, (640x480, 800x600, 1024x768,
        ' 1280x1024?). This also takes into account
        ' the size of the task bar, where ever it is.
        mptCurrentScreen.x = _
         GetSystemMetrics(SM_CXFULLSCREEN)
        mptCurrentScreen.y = _
         GetSystemMetrics(SM_CYFULLSCREEN)
        
        ' Get the pixels/inch ratio, as well.
        ptCurrentDPI.x = GetDeviceCaps(lngDC, LOGPIXELSX)
        ptCurrentDPI.y = GetDeviceCaps(lngDC, LOGPIXELSY)
        
        mptTwipsPerPixel.x = _
         adhcTwipsPerInch / ptCurrentDPI.x
        mptTwipsPerPixel.y = _
         adhcTwipsPerInch / ptCurrentDPI.y
    
        ' Release the information context.
        Call ReleaseDC(HWND_DESKTOP, lngDC)
    End If

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetScreenInfo", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Function GetIsPopup() As Boolean
    ' Is this form a popup form? To find that
    ' out, check the parent of the form. If
    ' it's Access, you know it's popup (otherwise,
    ' the parent is the MDI Client window.)

    Dim hWndParent As Long
    
    hWndParent = GetParent(frm.hwnd)
    GetIsPopup = (hWndParent = Application.hWndAccessApp)
End Function

Private Function GetIsSubform() As Boolean
    ' Is our form currently loaded as a subform?
    ' Check its Parent property to find out.
    
    Dim strName As String
    
    On Error Resume Next
    strName = frm.Parent.Name
    GetIsSubform = (Err.Number = 0)
    Err.Clear
End Function

Private Function ReferenceType() As ReferenceType
    Dim rt As ReferenceType
    
    ' Need some easy way to know what to
    ' use as a reference point in all
    ' position calculations. This function
    ' wraps that decision-making up.
    
    If IsPopup Then
        rt = rtPopup
    Else
        rt = rtNormal
    End If
    ReferenceType = rt
End Function

Private Function RctToCoords(rct As RECT) As COORDS
    ' Convert from a RECT struct to a COORDS struct.
    ' This seems to come up often.
    Dim C As COORDS
    With C
        .Left = rct.Left
        .Top = rct.Top
        .width = rct.Right - rct.Left
        .height = rct.Bottom - rct.Top
    End With
    RctToCoords = C
End Function

Private Sub GetClientOffsets()
    Dim p As POINTAPI
    Dim rct As RECT
    
    ' Get coords of upper-left corner of client area.
    ' Because Access coordinates are from the upper-left
    ' corner of the client area, but Windows coordinates
    ' measure from the upper-left corner of the main Access
    ' window, you need some way of calculating the offset
    ' of the MDI Client window. This procedure
    ' does that.
    
    ' Convert 0, 0 within the form's client area to
    ' absolute screen coordinates.
    p.x = 0
    p.y = 0
    Call ClientToScreen(frm.hwnd, p)
    Call GetWindowRect(frm.hwnd, rct)
    
    mptClientOffset.x = p.x - rct.Left
    mptClientOffset.y = p.y - rct.Top
End Sub

Private Sub GetParentCoords(hwnd As Long, C As COORDS)
    On Error GoTo HandleErrors
    Dim hWndParent As Long
    Dim rctParent As RECT
    
    Select Case ReferenceType
        Case rtNormal
            hWndParent = GetParent(hwnd)
            Call GetClientRect(hWndParent, rctParent)
        
        Case rtPopup
            ' Get the desktop coordinates.
            ' If you want to fill the ENTIRE Windows
            ' desktop, use this code.
            
            ' Call GetWindowRect(GetDesktopWindow(), rctParent)

            ' If you want to respect the task bar and other docked
            ' toolbars, use the following code, instead of the commented
            ' out code above.
            Call SystemParametersInfoRect( _
             SPI_GETWORKAREA, 0, rctParent, 0)
    End Select
    C = RctToCoords(rctParent)

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetParentCoords", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

