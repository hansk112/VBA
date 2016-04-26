VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' FormResize Class module

' A WithEvents reference to the form we're "ghosting."
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1

' Generally, set this to False. If you want to
' debug odd behaviors, change it to True.
#Const DEBUGGING = False

' ==================================
' Public Enums.
' ==================================
Public Enum ScaleControlsWhen
    scAtLoad = 1
    scYes = -1
    scNo = 0
End Enum

' ==================================
' Windows API declarations.
' ==================================

Private Declare Function SystemParametersInfoRect _
 Lib "user32" Alias "SystemParametersInfoA" _
 (ByVal uAction As Long, ByVal uParam As Long, _
 ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
 
Private Declare Function IsIconic _
 Lib "user32" _
 (ByVal hwnd As Long) As Long
 
 Private Declare Function IsZoomed _
  Lib "user32" _
  (ByVal hwnd As Long) As Long

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

Private Declare Function ShowWindow _
 Lib "user32" _
 (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

' Windows 95/98/NT4/2000 puts a 2-pixel
' border around the MDI client area, which
' doesn't get taken into account automatically.
' If you're using NT 3.51, you're on your own.
Private Const adhcBorderWidthX = 2
Private Const adhcBorderWidthY = 2

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1

Private Const HWND_DESKTOP = 0
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const SPI_GETWORKAREA = 48

Private Type POINTAPI
    x As Long
    y As Long
End Type

' Store rectangle coordinates.
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type COORDS
    Left As Long
    Top As Long
    width As Long
    height As Long
End Type

' ==================================
' Private Enums/Types
' ==================================

Private Type typDimensions
    Left As Long
    Top As Long
    width As Long
    height As Long
    ctlr As ControlResize
End Type

Private Enum LocationType
    ltUndefined = 0
    ltInControls = 1
    ltInGroups = 2
End Enum
    
' ==================================
' Simple properties
' ==================================

Public Parent As ControlResize
Public ParentControl As SubForm
' Leave fonts unscaled?
Public ScaleFonts As Boolean
' Leave original form size?
Public ScaleForm As Boolean
' Scale combo/listbox columns?
Public ScaleColumns As Boolean
' Scale controls? At load, always, or none.
Public ScaleControls As ScaleControlsWhen

' Min and Max Coordinates, in Twips.
Private msngMinWidth As Single
Private msngMinHeight As Single
Private msngMaxWidth As Single
Private msngMaxHeight As Single

Private mcolControls As Collection
Private mcolSections As Collection

Private mrct As RECT
Private mrctWindow As RECT
Private mcParent As COORDS
Private mfInload As Boolean

' ==================================
' Constants
' ==================================

Private Const acDatasheetView = 2

' Error constants
Private Const adhcErrDivisionByZero = 11
Private Const adhcErrInvalidProperty = 438
Private Const adhcErrTooLarge = 2100
Private Const adhcErrReferenceLost = 2467

Private Const adhcTwipsPerInch = 1440
' The maximum size for a form is 22 inches.
Private Const adhcMaxTwips = 22 * adhcTwipsPerInch

' ==================================
' Private variables
' ==================================

Private mfIsSubform As Boolean
Private mintInResize As Integer

Private mptCurrentScreen As POINTAPI
Private mptCurrentDPI As POINTAPI
Private mptDesignScreen As POINTAPI
Private mptDesignDPI As POINTAPI
Private mptTwipsPerPixel As POINTAPI

Private mlngWidthInTwips As Long
Private mlngHeightInTwips As Long

' ==================================
' Event Procedures
' ==================================

Private Sub Class_Initialize()
    ' Assume you want to scale fonts.
    ' You can override this by setting
    ' the ScaleFonts property manually.
    ScaleFonts = True
   
    ' Assume you want to scale the form.
    ' You can override this by setting
    ' the ScaleForm property manually.
    ScaleForm = True
    
    ' Assume you want to scale controls (that's the
    ' whole point, right?) You can turn off
    ' this behavior by setting this property to
    ' ScaleNo at any time.
    Me.ScaleControls = scYes
    
    
    ' Assume you want to scale list/combobox
    ' columns. If you don't, turn off this
    ' property.
    ScaleColumns = True
    
    ' Although these values will be
    ' 0 if you don't set them, it never
    ' hurts, and makes it clear. 0 is
    ' an invalid value for each of these.
    mptDesignScreen.x = 0
    mptDesignScreen.y = 0
    mptDesignDPI.x = 0
    mptDesignDPI.y = 0
    
    ' Just in case something goes
    ' wrong, you don't want to be
    ' dividing by 0 when you use
    ' these values.
    mptCurrentDPI.x = 15
    mptCurrentDPI.y = 15
    
    ' Initialize the resize "level" counter to 0.
    mintInResize = 0
    
    ' Unless told otherwise, set the
    ' maximum size to the size of the screen.
    msngMaxWidth = 1
    msngMaxHeight = 1
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mcolControls = Nothing
    Set mcolSections = Nothing
    Call Parent.CleanUp
End Sub

Private Sub frm_Load()
    mfInload = True
    If ScaleForm Then
        Call ScaleFormSize
        mfInload = False
    End If
End Sub

Private Sub frm_Resize()
    ' Handle the resize event for the form.
    ' This occurs after the form's Resize
    ' event procedure occurs.
    On Error GoTo HandleErrors
    If Not mfIsSubform Then
        ' No point doing this, from the
        ' Resize event, if this is a subform.
        ' Normally this won't happen, but if
        ' the programmer added the FormResize
        ' hooks into a form, and then uses
        ' it as a subform, you really
        ' don't want to be resizing twice.
        Call RescaleForm
    End If
    
    ' If you're not scaling the form,
    ' you don't want to set this flag
    ' to be false until you've resized
    ' the form once, from the form's
    ' first Resize event.
    mfInload = False
ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.Resize", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

' ==================================
' Public Methods and Properties
' ==================================

Public Property Get TwipsPerPixelX() As Long
    TwipsPerPixelX = mptTwipsPerPixel.x
End Property

Public Property Get TwipsPerPixelY() As Long
    TwipsPerPixelY = mptTwipsPerPixel.y
End Property

Public Property Get IsMaximized() As Boolean
    On Error GoTo HandleErrors
    IsMaximized = IsZoomed(frm.hwnd)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.IsMaximized", _
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
            Call HandleError("FormResize.IsMaximized", _
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
            Call HandleError("FormResize.IsMinimized", _
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
            Call HandleError("FormResize.IsMinimized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get MinWidth() As Single
    MinWidth = msngMinWidth
End Property

Public Property Get MaxWidth() As Single
    MaxWidth = msngMaxWidth
End Property

Public Property Get MinHeight() As Single
    MinHeight = msngMinHeight
End Property

Public Property Get MaxHeight() As Single
    MaxHeight = msngMaxHeight
End Property

Public Property Let MinWidth(value As Single)
    value = Max(value, 0)
    msngMinWidth = value
End Property

Public Property Let MinHeight(value As Single)
    value = Max(value, 0)
    msngMinHeight = value
End Property

Public Property Let MaxWidth(value As Single)
    value = Max(value, 0)
    msngMaxWidth = value
End Property

Public Property Let MaxHeight(value As Single)
    value = Max(value, 0)
    msngMaxHeight = value
End Property

Public Sub SetDesignCoords( _
 width As Long, height As Long, _
 DPIX As Long, DPIY As Long)
    
    ' Tell the code the original design-time
    ' coordinates of the screen, and the
    ' dots per inch value.
    
    ' Calling this method is optional.
    
    ' If you never set these values, they'll
    ' all be 0, and the later code will use
    ' the current screen values instead.
    mptDesignScreen.x = width
    mptDesignScreen.y = height
    mptDesignDPI.x = DPIX
    mptDesignDPI.y = DPIY
End Sub

Public Property Set Form(value As Form)
    On Error GoTo HandleErrors
    
    Call GetScreenInfo
    
    Set frm = value
    mfIsSubform = IsSubForm()
    If Not mfIsSubform Then
        ' If this isn't a subform,
        ' then it must be the main form.
        ' Set everything up here.
        frm.OnResize = "[Event Procedure]"
        frm.OnLoad = "[Event Procedure]"
    End If
    
    Call GetSize
    Call GetSectionInfo
    Call GetControlInfo
    
ExitHere:
    Exit Property
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.Form", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get Form() As Form
    Set Form = frm
End Property

Public Property Get Controls() As Collection
    Set Controls = mcolControls
End Property

Public Sub RescaleForm()
Attribute RescaleForm.VB_Description = "Forces a resize of all the controls on the form. Normally, you won't call this directly, but you can."

    ' Called from the Resize event of forms.
    ' Attempt to resize the form and all its
    ' controls.  Don't do anything if the
    ' current height of the form is very small,
    ' or if it's iconic.
    
    ' You can call this method directly to force
    ' a recalculation of the form control sizes.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    Dim rctNew As RECT
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim decFactorX As Variant
    Dim decFactorY As Variant
    
    On Error GoTo HandleErrors
    
    ' If this is a recursive resize call,
    ' just get out now.
    If mintInResize > 0 Then
        Exit Sub
    End If
    mintInResize = mintInResize + 1
    
    DoCmd.HourGlass True
    frm.Painting = False
    
    ' Make sure the user hasn't sized this thing down
    ' to the nubs.  If the client area is very short,
    ' it's time to call it quits.
    If IsIconic(frm.hwnd) Then
        ' Nothing to do, so exit now.
        GoTo ExitHere
    End If
    
    ' Check the height and width.
    ' Make sure they're within the
    ' Min/Max Width/Height values.
    Call CheckAndFixCoords
    
    Call GetClientRect(frm.hwnd, rctNew)
    ' Get the current width and height.
    ' If it's really small, get out.
    lngHeight = (rctNew.Bottom - rctNew.Top)
    If lngHeight < 10 Then
        GoTo ExitHere
    End If
    
    lngWidth = (rctNew.Right - rctNew.Left)
    
    decFactorX = CDec(lngWidth / (mrct.Right - mrct.Left))
    decFactorY = CDec(lngHeight / (mrct.Bottom - mrct.Top))
        
FactorOK:
    ' If you actually want to do some resizing, do it now.
    Call ScaleFormContents(decFactorX, decFactorY)

ExitHere:
    On Error Resume Next
    mintInResize = mintInResize - 1
    frm.Painting = True
    DoCmd.HourGlass False
    Exit Sub
    

HandleErrors:
    Select Case Err.Number
        Case adhcErrInvalidProperty
            Resume ExitHere
            
        Case adhcErrDivisionByZero
            ' If you get a division by zero error,
            ' simply set both factors to be 1 (the
            ' original size) and let it go.
            decFactorX = CDec(1)
            decFactorY = CDec(1)
            Resume FactorOK
            
        Case Else
            Call HandleError("FormResize.RescaleForm", Err.Number, Err.Description)
            Resume Next
    End Select
End Sub

Public Sub ScaleFormContents( _
 decFactorX As Variant, decFactorY As Variant)

    ' Actually do the work to resize all the controls
    ' on the given form, and then resize the form
    ' itself.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
        
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim intI As Integer
    Dim intGroups As Integer
    Dim aContainers() As typDimensions
    Dim lt As LocationType
    Dim decFactor As Variant
    Dim decFontSize As Variant
    Dim ctlr As ControlResize
    Dim ctl As control
    Dim frm As Form
    Dim rct As RECT
    Dim CC As COORDS
    
    On Error GoTo HandleErrors

TryAgain:
    Set frm = Me.Form
    ' If this is a datasheet form, just
    ' get out now. We tried, really, to
    ' handle datasheets, but Access just
    ' doesn't give you the information you'd
    ' need to do it correctly. Try finding
    ' out the default row height, for example!
    If frm.CurrentView = acDatasheetView Then
        GoTo ExitHere
    End If
    
    ' Keep track of where you are in the code,
    ' so the error handler can do the right thing.
    lt = ltUndefined

    ' Go through and deal with all the groups first.
    intGroups = GetContainerInfo(aContainers(), decFactorX, decFactorY)

    ' Now go back and deal with all the rest of the controls.
    lt = ltInControls
    
    For Each ctlr In mcolControls
        Set ctl = ctlr.control
                
        ' Set the coordinates.
        
        Select Case ctl.ControlType
            Case acOptionGroup, acTabCtl, acPage
                GoTo NextCtl
        End Select
    
        If FloatThisControl(ctlr) Then
            CC = FloatTheControl(ctlr)
            ctl.Left = CC.Left
            ctl.Top = CC.Top
        End If
        
        If StretchThisControl(ctlr) Then
            CC = StretchTheControl(ctlr)
            If CC.width > 0 Then
                ctl.width = CC.width
            End If
            If CC.height > 0 Then
                ctl.height = CC.height
            End If
        End If
        
        If ScaleThisControl(ctlr) Then
            ' So the control isn't a container.
            ' Therefore, just scale it to the correct size.

            CC = ScaleTheControl(ctlr, decFactorX, decFactorY)
            With ctl
                .Top = CC.Top
                .Left = CC.Left
                .width = CC.width
                .height = CC.height
            End With
                    
            If ctlr.HasColumns And ScaleColumns Then
                ctl.ColumnWidths = ctlr.ScaledColumnWidths(decFactorX)
            End If
            
            Call ScaleTheFont(ctlr, decFactorX, decFactorY)
        End If
        
NextCtl:
        ' If you've got a subform control,
        ' call the whole thing recursively.
        ' Luckily, this can't go more than 2 levels deep.
        If ctl.ControlType = acSubform Then
            If ScaleThisForm Then
                Call ctlr.frm.ScaleFormContents(decFactorX, decFactorY)
            End If
        End If
    Next ctlr

    ' Go through and fix up the option groups/subforms,
    ' which may have been distorted by changes to
    ' the internal controls.
        
    lt = ltInGroups
    For intI = intGroups To 1 Step -1
        Set ctl = aContainers(intI).ctlr.control
        With aContainers(intI)
            ctl.Top = .Top
            ctl.Left = .Left
            ctl.width = .width
            ctl.height = .height
        End With
        Call ScaleTheFont(aContainers(intI).ctlr, decFactorX, decFactorY)
    Next intI
    If ScaleThisForm() Then
        Call SetFormCoords(decFactorX, decFactorY)
    End If

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case adhcErrReferenceLost
            ' If you've got a lost reference,
            ' go try to reset them.
            If ResetReferences(Me) Then
                Resume TryAgain
            Else
                Resume ExitHere
            End If
        Case adhcErrInvalidProperty
            ' Invalid property? Keep going.
            Resume Next
        Case adhcErrTooLarge
            Select Case lt
                ' If the control's too large, you need to
                ' increase the height of the section. Make
                ' it a tiny bit larger, until this error doesn't occur.
                Case ltInControls
                    Call IncreaseHeight(frm.Section(ctl.Section))
                    Resume
                Case ltInGroups
                    With aContainers(intI).ctlr.control
                        If TypeOf .Parent Is Page Then
                            Call IncreaseHeight(.Parent)
                        Else
                            Call IncreaseHeight(.Parent.Section(.Section))
                        End If
                    End With
                    Resume
                Case ltUndefined
                    Resume Next
            End Select
        Case Else
            Call HandleError("FormResize.SetSize", Err.Number, Err.Description)
            Resume Next
    End Select
End Sub

Public Property Set FormReset(value As Form)
    ' You should never need to call this method.
    
    ' Simply set the form reference. This property
    ' is used when resetting references, should
    ' Access lose a reference to a control.
    Set frm = value
End Property

Public Property Get WidthInTwips() As Long
    WidthInTwips = mlngWidthInTwips
End Property

Public Property Get HeightInTwips() As Long
    HeightInTwips = mlngHeightInTwips
End Property

Public Sub CleanUp()
    ' Release references to
    ' other objects.
    Set Parent = Nothing
End Sub

' ==================================
' Private Methods
' ==================================

Private Function IsSubForm() As Boolean
    ' Is our form currently loaded as a subform?
    ' Check its Parent property to find out.
    
    Dim strName As String
    
    On Error Resume Next
    strName = frm.Parent.Name
    IsSubForm = (Err.Number = 0)
    Err.Clear
End Function

Private Sub GetControlInfo()
    ' This procedure counts on the module-level
    ' collection, mcolControls.
    
    Dim ctl As control
    Dim ctlr As ControlResize
    
    On Error GoTo HandleErrors
    
    If Not (mcolControls Is Nothing) Then
        Exit Sub
    End If
    
    Set mcolControls = New Collection
    For Each ctl In frm.Controls
        Set ctlr = New ControlResize
        With ctlr
            Set .Parent = Me
            Set .control = ctl
            .Name = ctl.Name
            
            .Top = ctl.Top
            .Left = ctl.Left
            .width = ctl.width
            .height = ctl.height
            .FontSize = ctl.FontSize
            ' Most controls won't have a ColumnWidths
            ' property, but ListBox and ComboBox controls do.
            .ColumnWidths = ctl.ColumnWidths
        End With
        ' Each form (and form within subforms)
        ' keeps track of its own controls.
        ' That way, names are guaranteed
        ' to be a unique key value.
        mcolControls.Add ctlr, ctlr.Name
    Next ctl

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case adhcErrInvalidProperty
            Resume Next
        Case Else
            Call HandleError("FormResize.GetControlInfo", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Sub GetSize()

    ' Fill in rct with the coordinates of the client area.
    ' Fill in rctWindow with coordinates of the window.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    Dim hWndParent As Long
    Dim rctParent As RECT

    On Error GoTo HandleErrors

    ' Find the position of the window in question, in
    ' relation to its parent window (the Access desktop,
    ' the MDIClient window).
    hWndParent = GetParent(frm.hwnd)

   ' Get the coordinates of the current window and its parent.
    Call GetClientRect(frm.hwnd, mrct)
    Call GetWindowRect(frm.hwnd, mrctWindow)

    ' Catch the case where the form is Popup (that is,
    ' its parent is NOT the Access main window.)  In that
    ' case, don't subtract off the coordinates of the
    ' Access MDIClient window.
    If hWndParent <> Application.hWndAccessApp Then
        Call GetWindowRect(hWndParent, rctParent)

        ' Subtract off the left and top parent coordinates, since you
        ' need coordinates relative to the parent for the MoveWindow()
        ' function call.
        With mrctWindow
            .Left = .Left - rctParent.Left
            .Top = .Top - rctParent.Top
            .Right = .Right - rctParent.Left
            .Bottom = .Bottom - rctParent.Top
        End With
    End If
    With mrct
        mlngHeightInTwips = (.Bottom - .Top) * mptTwipsPerPixel.y
        mlngWidthInTwips = (.Right - .Left) * mptTwipsPerPixel.x
    End With
ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.GetSize", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Sub GetSectionInfo()
    ' Get information about each of
    ' the five possible sections.
    
    Dim secR As SectionResize
    Dim i As Integer
    
    On Error GoTo HandleErrors
    
    Set mcolSections = New Collection
    
    ' There are 5 possible sections in a form,
    ' but they might not all be there.
    On Error Resume Next
    
    For i = 0 To 4
        Set secR = New SectionResize
        secR.height = frm.Section(i).height
        secR.Section = i
        
        ' If you use a numeric key in a collection,
        ' it's possible that retrieving items may
        ' be confusing (is it an index, or a key?)
        ' Therefore, tack a leading "S" onto the
        ' item number, so it's clear that you're
        ' referring to the key, not the index.
        mcolSections.Add secR, "S" & i
    Next i

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.GetSectionInfo", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Sub MoveForm(decFactorX As Variant, decFactorY As Variant)
    Dim intWidth As Integer
    Dim intHeight As Integer
    
    ' Move the form. Only called when the form is first
    ' loaded, if it needs to be moved to match the
    ' original positioning.
    
    On Error GoTo HandleErrors
    With mrctWindow
        intWidth = Int((.Right - .Left) * decFactorX)
        intHeight = Int((.Bottom - .Top) * decFactorY)

        .Left = Int(.Left * decFactorX)
        .Top = Int(.Top * decFactorY)
        .Right = .Left + intWidth
        .Bottom = .Top + intHeight

        mintInResize = mintInResize + 1
        Call MoveWindow(frm.hwnd, .Left, .Top, _
         intWidth, intHeight, bRepaint:=1)
        
        ' Give Windows a chance to catch up...
        DoEvents
        ' Use these flags to make sure you
        ' don't call the from resizing code
        ' recursively. There's a check in
        ' that code that exits the procedure
        ' if mintInResize > 0.
        On Error Resume Next
        
        DoCmd.RunCommand acCmdSizeToFitForm
        mintInResize = mintInResize - 1

    End With
    
ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.MoveForm", Err.Number, Err.Description)
            Resume ExitHere
    End Select
End Sub

Private Sub FixSections(decFactorY As Variant)

    ' Loop through all the sections of the form,
    ' up to 5 sections, setting the height of
    ' each.  If a section isn't there, just keep
    ' on going.

    Dim i As Integer

    ' There are 5 possible sections in a form,
    ' but they might not all be there.
     On Error Resume Next
    With frm
        For i = 0 To 4
            .Section(i).height = _
             Min(RoundWhole(CDec(mcolSections("S" & i).height * decFactorY)), adhcMaxTwips)
        Next i
    End With
End Sub

Private Function Min(varValue1 As Variant, varValue2 As Variant) As Variant
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    If varValue1 < varValue2 Then
        Min = varValue1
    Else
        Min = varValue2
    End If
End Function

Private Function Max(varValue1 As Variant, varValue2 As Variant) As Variant
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    If varValue1 > varValue2 Then
        Max = varValue1
    Else
        Max = varValue2
    End If
End Function

Private Function GetContainerInfo(aContainers() As typDimensions, _
 decFactorX As Variant, decFactorY As Variant)

    ' Store away information about controls that
    ' contain other controls (frames/tabs).
    ' Return the number of these controls that
    ' were found.

    Dim fDoIt As Boolean
    Dim intGroups As Integer
    Dim ctl As control
    Dim ctlr As ControlResize
    Dim CC As COORDS

    On Error GoTo HandleErrors
    intGroups = 0
    
    For Each ctlr In mcolControls
TryAgain:
        ' Get at the control referenced by ctlR.
        Set ctl = ctlr.control
        Select Case ctl.ControlType
            ' If the control's an OptionGroup or a
            ' Tab control, store away the coordinates
            ' of the control.
            Case acOptionGroup, acTabCtl
                intGroups = intGroups + 1
                ReDim Preserve aContainers(1 To intGroups)
                
                If ScaleThisControl(ctlr) Then
                    CC = ScaleTheControl(ctlr, decFactorX, decFactorY)
                ElseIf FloatThisControl(ctlr) Then
                    CC = FloatTheControl(ctlr)
                ElseIf StretchThisControl(ctlr) Then
                    CC = StretchTheControl(ctlr)
                Else
                    CC.Left = ctl.Left
                    CC.Top = ctl.Top
                    CC.width = ctl.width
                    CC.height = ctl.height
                End If
                With aContainers(intGroups)
                    Set .ctlr = ctlr
                    .Left = CC.Left
                    .Top = CC.Top
                    .width = CC.width
                    .height = CC.height
                End With
        End Select
        
NextControl:
    Next ctlr
    
ExitHere:
    GetContainerInfo = intGroups
    Exit Function
    
HandleErrors:
    Select Case Err.Number
        Case adhcErrInvalidProperty
            Resume Next
        Case adhcErrReferenceLost
            If ResetReferences(Me) Then
                Resume TryAgain
            Else
                Resume ExitHere
            End If
    End Select
    Resume NextControl
End Function

Private Function ChangeFont(ctl As control) As Boolean

    ' Decide whether or not to change the font,
    ' based on the control type.

    ChangeFont = False
    Select Case ctl.ControlType
        Case acTextBox, acComboBox, acListBox, acLabel, _
         acCommandButton, acToggleButton, acTabCtl
            ChangeFont = True
        Case Else
    End Select
End Function

Private Function RoundWhole(varValue As Variant) As Variant
    ' Round to the nearest whole number.
    RoundWhole = Int(CDec(varValue) + 0.5)
End Function

Private Sub IncreaseHeight(sec As Object)
    On Error Resume Next
    sec.height = Int(sec.height * 1.1)
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

Private Function ResetReferences( _
 frmR As FormResize, _
 Optional Level As Integer = 1) As Boolean
    ' Only here in case of emergency. If you scale
    ' the form too small, Access throws away subforms,
    ' and recreates them when the form gets larger.
    ' This causes the code undue anguish, as it's attempting
    ' to maintain pointers to all the objects on the form.
    
    ' This procedure attempts to rebuild all the
    ' necessary references. It may fail. If so, the
    ' next time it gets called, perhaps it will
    ' succeed. If it never succeeds, well, then your
    ' form will look very odd. But it's always succeeded
    ' sooner or later, in tests, depending on how you resize
    ' your form.
    
    Dim ctl As control
    Dim ctlr As ControlResize
    Dim frmLocal As Form
    
    On Error GoTo HandleErrors
    
    Set frmLocal = frmR.ParentControl.Form
    If Level = 1 Then
        ' Only reset the form-level frm variable
        ' at the top level subform.
        Set frm = frmLocal
    End If
    
    For Each ctlr In frmR.Controls
        Set ctl = frmLocal.Controls(ctlr.Name)
        Set ctlr.ResetControl = ctl
        Set ctlr.Parent = frmR
        If ctl.ControlType = acSubform Then
            Set ctlr.frm.ParentControl = ctl
            Call ResetReferences(ctlr.frm, Level + 1)
        End If
    Next ctlr
    ResetReferences = True
    
ExitHere:
    Exit Function

HandleErrors:
    ResetReferences = False
    Select Case Err.Number
        Case adhcErrInvalidProperty
            Resume Next
        Case Else
            Call HandleError("FormResize.ResetReferences", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Function

Private Function ScaleFormSize()

    ' Called from the Open event of forms.
    ' Attempts to scale the form appropriately
    ' for the given screen size, as compared
    ' to the size screen on which it was designed.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    

    Dim decFactorX As Variant
    Dim decFactorY As Variant

    On Error GoTo HandleErrors
    
    Call GetScreenScale(decFactorX, decFactorY)
    
    ' If the x and y factors are both 1, there's nothing
    ' to do.
    
    ' If you don't want forms to expand (they were created on a
    ' lower-resolution device than the current device), but only
    ' shrink (they were created on a higher-resolution device
    ' than the current device), then use the next line instead
    ' of the current If...Then line.
    'If (decFactorX < 1) Or (decFactorY < 1) Then
    
    If (decFactorX <> 1) Or (decFactorY <> 1) Then
    
        Call ScaleFormContents(decFactorX, decFactorY)
        Call MoveForm(decFactorX, decFactorY)
        
    End If
    
ExitHere:
    Exit Function

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.ScaleFormSize", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Function

Private Sub GetScreenScale( _
 ByRef decFactorX As Variant, ByRef decFactorY As Variant)

    ' Set the module-level variables that track
    ' the current screen scaling. Pass the two
    ' scaling factors back as ByRef parameters.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors

    ' If you don't know now what the
    ' design-time coordinates were, simply
    ' use the current ones instead.
    If mptDesignScreen.x = 0 Then
        mptDesignScreen.x = mptCurrentScreen.x
    End If
    If mptDesignScreen.y = 0 Then
        mptDesignScreen.y = mptCurrentScreen.y
    End If
    
    If mptDesignDPI.x = 0 Then
        mptDesignDPI.x = mptCurrentDPI.x
    End If
    
    If mptDesignDPI.y = 0 Then
        mptDesignDPI.y = mptCurrentDPI.y
    End If
        
    ' Get the ratio of the current screen size to the design-time
    ' screen size.
    ' Take into account the differences in the display
    ' resolutions.  At 640x480, you get more twips per pixel (15)
    ' as opposed to 12 at higher resolutions.
    decFactorX = CDec(mptCurrentScreen.x / mptDesignScreen.x) _
     * (mptDesignDPI.x / mptCurrentDPI.x)
    
    decFactorY = CDec(mptCurrentScreen.y / mptDesignScreen.y) _
     * (mptDesignDPI.y / mptCurrentDPI.y)

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case adhcErrDivisionByZero
            ' It seems that the first time you call
            ' GetDeviceCaps under Win95 after you've done
            ' a quick change on the resolution, it returns 0
            ' for the screen size.  This will hopefully correct
            ' that problem.
            Resume
        Case Else
            Call HandleError("FormResize.GetScreenScale", Err.Number, Err.Description)
            Resume ExitHere
    End Select
End Sub

Private Sub GetScreenInfo()
    ' Get the information context you need to find the screen info.
    ' This procedure fills in the module variables:
    ' mptCurrentScreen.x, mptCurrentScreen.y
    ' mintmptCurrentDPI.x, mintmptCurrentDPI.y
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    Dim rctParent As RECT
    Dim lngDC As Long
    Dim ctemp As COORDS
    
    On Error GoTo HandleErrors
    
    lngDC = GetDC(HWND_DESKTOP)

    ' If the call to GetDC didn't fail (and it had
    ' better not, or things are really busted),
    ' then get the info.
    
    If lngDC <> 0 Then
        ' Find the number of pixels in both directions on the
        ' screen, (640x480, 800x600, 1024x768, 1280x1024?). This
        ' also takes into account the size of the task bar, where
        ' ever it is.
        ' Get the desktop coordinates.
        ' If you want to fill the ENTIRE Windows
        ' desktop, use this code.
        
        ' Call GetWindowRect(GetDesktopWindow(), rctParent)

        ' If you want to respect the task bar and other docked
        ' toolbars, use the following code, instead of the commented
        ' out code above.
        Call SystemParametersInfoRect( _
         SPI_GETWORKAREA, 0, rctParent, 0)
        ctemp = RctToCoords(rctParent)
                 
        mptCurrentScreen.x = ctemp.width
        mptCurrentScreen.y = ctemp.height
        
        ' Set up the mcParent structure
        mcParent.width = mptCurrentScreen.x
        mcParent.height = mptCurrentScreen.y
        mcParent.Top = 0
        mcParent.Left = 0
        
        ' Get the pixels/inch ratio, as well.
        mptCurrentDPI.x = GetDeviceCaps(lngDC, LOGPIXELSX)
        mptCurrentDPI.y = GetDeviceCaps(lngDC, LOGPIXELSY)
        
        mptTwipsPerPixel.x = adhcTwipsPerInch / mptCurrentDPI.x
        mptTwipsPerPixel.y = adhcTwipsPerInch / mptCurrentDPI.y

        ' Release the information context.
        Call ReleaseDC(HWND_DESKTOP, lngDC)
    End If

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.GetScreenInfo", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Function ScaleThisControl(ctlr As ControlResize) As Boolean
    ' Based on the form's ScaleControls property,
    ' and the control's ScaleControl property, should you
    ' resize the control?
    
    ' If the control says No, then don't scale.
    ' If the control says Yes, or the control says use
    ' the default value, and the default value is True,
    ' then scale.
    
    Dim fDoIt As Boolean
    
    Select Case ctlr.ScaleIt
        Case csNo
            fDoIt = False
        Case csYes
            fDoIt = True
        Case csDefault
            Select Case ScaleControls
                Case scNo
                    fDoIt = False
                Case scYes
                    fDoIt = True
                Case scAtLoad
                    fDoIt = mfInload
            End Select
    End Select
    ScaleThisControl = fDoIt
End Function
    
Private Function FloatThisControl(ctlr As ControlResize) As Boolean
    ' Based on the control's FloatIt property, should
    ' this control be floated?
    FloatThisControl = (ctlr.FloatIt <> cfNone)
End Function

Private Function StretchThisControl(ctlr As ControlResize) As Boolean
    ' Based on the control's SizeIt property, should
    ' this control be sized?
    StretchThisControl = (ctlr.SizeIt <> czNone)
End Function

Private Sub SetFormCoords(decFactorX As Variant, decFactorY As Variant)
    ' Fix up all the section heights.
    mintInResize = mintInResize + 1
    Call FixSections(decFactorY)
    If mfIsSubform Then
        frm.width = ParentControl.width
    Else
        frm.width = Min(RoundWhole(mlngWidthInTwips * decFactorX), adhcMaxTwips)
    End If
    mintInResize = mintInResize - 1
End Sub

Private Function FloatTheControl(ctlr As ControlResize) As COORDS
    Dim flt As ControlFloat
    Dim ctl As control
    Dim CC As COORDS
    
    On Error Resume Next
    Set ctl = ctlr.control
    flt = ctlr.FloatIt
    CC.width = ctl.width
    CC.height = ctl.height
    If (flt And cfRight) Then
        CC.Left = frm.InsideWidth - ctlr.OffsetFromRight
    Else
        CC.Left = ctl.Left
    End If
    If (flt And cfBottom) Then
        CC.Top = frm.InsideHeight - ctlr.OffsetFromBottom
    Else
        CC.Top = ctl.Top
    End If
    FloatTheControl = CC
    Set ctl = Nothing
    Err.Clear
End Function

Private Function StretchTheControl(ctlr As ControlResize) As COORDS
    Dim sz As ControlSize
    Dim ctl As control
    Dim CC As COORDS
    
    On Error Resume Next
    sz = ctlr.SizeIt
    Set ctl = ctlr.control
    CC.Top = ctl.Top
    CC.Left = ctl.Left
    If (sz And czRight) Then
        CC.width = frm.InsideWidth - (ctlr.Left + ctlr.OffsetFromRight - ctlr.OriginalWidth)
    Else
        CC.width = ctl.width
    End If
    If (sz And czBottom) Then
        CC.height = frm.InsideHeight - (ctlr.Top + ctlr.OffsetFromBottom - ctlr.OriginalHeight)
    Else
        CC.height = ctl.height
    End If
    StretchTheControl = CC
    Set ctl = Nothing
    Err.Clear
End Function

Private Function ScaleTheControl(ctlr As ControlResize, _
 decFactorX As Variant, decFactorY As Variant) As COORDS
    Dim ctl As control
    Dim CC As COORDS
    
    On Error Resume Next
    With CC
        .Top = Int(CDec(ctlr.Top * decFactorY))
        .Left = Int(CDec(ctlr.Left * decFactorX))
        .width = Int(CDec(ctlr.width * decFactorX))
        .height = Int(CDec(ctlr.height * decFactorY))
    End With
    ScaleTheControl = CC
    Err.Clear
End Function

Private Function ScaleThisForm() As Boolean
    ScaleThisForm = (ScaleControls = scYes) Or _
     ((ScaleControls = scAtLoad) And mfInload)
End Function

Private Sub ScaleTheFont( _
 ctlr As ControlResize, decFactorX As Variant, decFactorY As Variant)
    Dim ctl As control
    Dim decFactor As Variant
    Dim decFontSize As Variant
    
    On Error GoTo HandleErrors
    
    If ScaleFonts Then
        Set ctl = ctlr.control
        ' Fix up the font, if this control has a font
        ' that needs to be fixed up.
        
        ' For font changes, use the minimum of the
        ' X and Y scaling factors.
        decFactor = CDec(Min(decFactorX, decFactorY))
        If ChangeFont(ctl) Then
            decFontSize = RoundWhole(CDec(ctlr.FontSize * decFactor))
            ' The minimum font size is 1. The maximum size is 127.
            ' Rather than call the Min and max functions here,
            ' we'll simply inline the code. This gets called
            ' so often, you hate to pay the price
            ' of calling a two extra functions for each control.
            If decFontSize < 1 Then
                decFontSize = 1
            End If
            If decFontSize > 127 Then
                decFontSize = 127
            End If
            ctl.FontSize = decFontSize
        End If
    End If
ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.ScaleTheFont", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Sub SetSize( _
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
    
    Dim C As COORDS
    
    C = GetCoords()
    With C
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
            Call HandleError("FormResize.SetSize", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Function GetCoords() As COORDS

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors
    Dim hWndParent As Long
    Dim rctParent As RECT
    Dim rct As RECT
    Dim C As COORDS

    ' Find the position of the window in question, in
    ' relation to its parent window (the Access desktop,
    ' the MDIClient window).
    hWndParent = GetParent(frm.hwnd)

   ' Get the coordinates of the current window and its parent.
    Call GetWindowRect(frm.hwnd, rct)

    ' Catch the case where the form is Popup (that is,
    ' its parent is NOT the Access main window.)  In that
    ' case, don't subtract off the coordinates of the
    ' Access MDIClient window.
    If hWndParent <> Application.hWndAccessApp Then
        ' Get the MDI Client window's coordinates.
        Call GetWindowRect(hWndParent, rctParent)
        mcParent = RctToCoords(rctParent)

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
    End If
    C = RctToCoords(rct)

ExitHere:
    GetCoords = C
    Exit Function

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.GetCoords", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
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

Private Sub CheckAndFixCoords()
    ' Make sure the form coordinates fit within
    ' the specified range.
    Dim C As COORDS
    Dim sngNew As Single
    Dim fChanged As Boolean
    Dim lngCompare As Long
    
    ' If the user wants to maximize
    ' this form, just let them.
    If IsMaximized() Then
        GoTo ExitHere
    End If
    
    On Error GoTo HandleErrors
    
    ' If any of the Min/Max coords is non-zero
    ' then get to work.
    
    If msngMinWidth + msngMaxWidth + msngMinHeight + msngMaxHeight <> 0 Then
        ' There's something to do.
        ' Get the coordinates, in pixels.
        C = GetCoords()
        
        ' If a property is >1 then the size indicates
        ' the min/max size IN TWIPS. If it's <=1 then
        ' the value indicates the percentage of the available size.
        If msngMinWidth > 0 Then
            If msngMinWidth <= 1 Then
                ' Use fraction of screen size
                lngCompare = msngMinWidth * mcParent.width
            Else
                lngCompare = msngMinWidth \ mptTwipsPerPixel.x
            End If
            sngNew = Max(C.width, lngCompare)
            If sngNew <> C.width Then
                C.width = sngNew
                fChanged = True
            End If
        End If
        If msngMaxWidth > 0 Then
            If msngMaxWidth <= 1 Then
                ' Use fraction of screen size
                lngCompare = msngMaxWidth * mcParent.width
            Else
                lngCompare = msngMaxWidth \ mptTwipsPerPixel.x
            End If
            sngNew = Min(C.width, lngCompare)
            If sngNew <> C.width Then
                C.width = sngNew
                fChanged = True
            End If
        End If
        If msngMinHeight > 0 Then
            If msngMinHeight <= 1 Then
                ' Use fraction of screen size
                lngCompare = msngMinHeight * mcParent.height
            Else
                lngCompare = msngMinHeight \ mptTwipsPerPixel.y
            End If
            sngNew = Max(C.height, lngCompare)
            If sngNew <> C.height Then
                C.height = sngNew
                fChanged = True
            End If
        End If
        If msngMaxHeight > 0 Then
            If msngMaxHeight <= 1 Then
                ' Use fraction of screen size
                lngCompare = msngMaxHeight * mcParent.height
            Else
                lngCompare = msngMaxHeight \ mptTwipsPerPixel.x
            End If
            sngNew = Min(C.height, lngCompare)
            If sngNew <> C.height Then
                C.height = sngNew
                fChanged = True
            End If
        End If
        If fChanged Then
            ' If the form is going to shrink because
            ' of this check, you'll need to turn off
            ' screen display, do the work, and then
            ' turn it on again, to force Access to
            ' repaint the MDI client area.
            DoCmd.Echo False
            Call SetSize(C.Left, C.Top, C.width, C.height)
            DoCmd.Echo True
        End If
    End If
    
ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormResize.CheckAndFixCoords", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

