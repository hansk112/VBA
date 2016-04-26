Attribute VB_Name = "Calendar"
Option Compare Database
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

Private Const adhcCalendarForm As String = "frmPopupCalendar"

Public Function adhDoCalendar( _
 Optional ByVal varStartDate As Variant) As Variant
     '
    ' This is the public entry point.
    ' If the passed in date is missing (as it will
    ' be if someone just opens the Calendar form
    ' raw), start on the current day.
    ' Otherwise, start with the date that is passed in.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    '
    Dim dtmStartDate As Date
    
    ' If they passed a value at all, attempt to
    ' use it as the start date.
    If IsMissing(varStartDate) Then
        dtmStartDate = Date
    Else
        If IsDate(varStartDate) Then
            dtmStartDate = varStartDate
        Else
            ' OK, so they passed a value that
            ' wasn't a date.
            ' Just use today's date in that case, too.
            dtmStartDate = Date
        End If
    End If
    DoCmd.OpenForm formname:=adhcCalendarForm, _
     windowmode:=acDialog, _
     OpenArgs:=dtmStartDate

    ' Stop here and wait.

    '
    ' If the form is still loaded, then get the
    ' final chosen date from the form.  If it isn't,
    ' return Null.
    If IsOpen(adhcCalendarForm) Then
        adhDoCalendar = Forms(adhcCalendarForm).value
        'MsgBox "Forms(adhcCalendarForm).Value is " & Forms(adhcCalendarForm).Value
        DoCmd.Close acForm, adhcCalendarForm
    Else
        adhDoCalendar = Null
    End If
End Function

Public Function adhDoCalendarLoop( _
 Optional ByVal varStartDate As Variant) As Variant
    '
    ' You may not want to use the acDialog flag.
    ' For example, if you want to set properties
    ' of the form as it's loading, you can't really
    ' use acDialog (unless you do some ugly OpenArgs
    ' handling. This example opens the form invisibly,
    ' allows you to modify properties, and then loops
    ' until you close or hide the form.
    
    ' This is the public entry point.
    ' If the passed in date is missing (as it will
    ' be if someone just opens the Calendar form
    ' raw), start on the current day.
    ' Otherwise, start with the date that is passed in.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    '
    ' If they passed a value at all, attempt to
    ' use it as the start date.
        
    Dim dtmStartDate As Date
    'Dim frm As Form_frmCalendar
    Dim frm As Form_frmPopupCalendar
    
        
    If IsMissing(varStartDate) Then
        dtmStartDate = Date
    Else
        If IsDate(varStartDate) Then
            dtmStartDate = varStartDate
        Else
            ' OK, so they passed a value that
            ' wasn't a date.
            ' Just use today's date in that case, too.
            dtmStartDate = Date
        End If
    End If
    
    ' Open the form module (effectively, loading an
    ' instance of the form, as well). Then, set
    ' properties of that form.
    Set frm = New Form_frmPopupCalendar
    frm.ShowOKCancel = False
    frm.value = dtmStartDate
    
    ' Call the ShowFormAndWait procedure. This
    ' returns True if the form is still loaded,
    ' but invisible, or False if it's been unloaded.
    If ShowFormAndWait(frm) Then
        ' If the form is still loaded, then get the
        ' final chosen date from the form.  If it isn't,
        ' return Null.
        adhDoCalendarLoop = frm.value
        DoCmd.Close acForm, adhcCalendarForm
    Else
        adhDoCalendarLoop = Null
    End If
End Function

Private Function IsOpen(strName As String, _
 Optional intObjectType As AcObjectType = acForm)
    ' Returns True if strName is open, False otherwise.
    ' Assume the caller wants to know about a form.
    IsOpen = (SysCmd(acSysCmdGetObjectState, _
     intObjectType, strName) <> 0)
End Function

Public Function ShowFormAndWait(frm As Form) As Boolean
    Dim blnCancelled As Boolean
    Dim lngLoop As Long
    Dim strName As String
    
    ' Take an opened form module, display the
    ' form, and loop until the user closes or hides
    ' the form. Using this technique, you can
    ' create popup forms without losing the capability
    ' of modifying properties of the form before
    ' displaying it.
    
    ' In:
    '     frm: An opened form class module reference
    ' Out:
    '     Return Value: True if the form was hidden, False if closed.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    '
    ' Note: In order for this procedure to work
    ' correctly, the form you've opened must have its
    ' Popup and MOdal (OK, Popup isn't required, but
    ' Modal is) set to True.
    
    ' Note: Because this procedure relies on the
    ' form's Name property, it will not work for
    ' forms that have multiple instances open.
    
    ' Check for the form closing or hiding
    ' every adhcInterval times through the loop.
    ' We've found 1000 to work fine, as a balance
    ' between responsiveness and not sucking the
    ' life out of your application. You may want
    ' to change this value to suit your own needs.
    Const adhcInterval As Long = 1000
    
    strName = frm.Name
    frm.Visible = True
        
    Do
        If lngLoop Mod adhcInterval Then
            DoEvents
            ' Is the form still open?
            If Not IsOpen(strName) Then
                blnCancelled = True
                Exit Do
            End If
            ' OK, it's still open. Is it visible?
            If Not frm.Visible Then
                blnCancelled = False
                Exit Do
            End If
            lngLoop = 0
        End If
        lngLoop = lngLoop + 1
    Loop
    ShowFormAndWait = Not blnCancelled
End Function

