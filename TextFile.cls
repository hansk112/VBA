VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All Rights Reserved.

Option Compare Database
Option Explicit

' Private variables for properties implemented
' using property procedures
Private mstrPath As String
Private mhFile As Long
Private mlngOpenMode As TextFileOpenMode
Private mfIsOpen As Boolean
Private mstrCurrentLine As String

' Enumeration for file open mode
Public Enum TextFileOpenMode
    tfOpenReadOnly
    tfOpenReadWrite
    tfOpenAppend
    tfOpenReadOnlyExists
End Enum

Public Function FileOpen() As Boolean
    On Error GoTo HandleError
    
    ' If a file is already open, close it
    If Me.IsOpen Then
        Me.FileClose
    End If
    
    ' Get next available file handle
    mhFile = FreeFile
    
    ' Open file based on file open mode property
    Select Case Me.OpenMode
        Case tfOpenReadOnly
            Open Me.Path For Binary Access Read As mhFile
        Case tfOpenReadWrite
            Open Me.Path For Binary Access Read Write As mhFile
        Case tfOpenAppend
            Open Me.Path For Append Access Read Write As mhFile
        Case tfOpenReadOnlyExists
            Open Me.Path For Input Access Read As mhFile
    End Select
    
    ' Set IsOpen property variable and return value
    mfIsOpen = True
    FileOpen = True
    
ExitProc:
    Exit Function

HandleError:
    FileOpen = False
    Resume ExitProc
End Function

Public Sub FileClose()
    Close Me.Handle
    mfIsOpen = False
End Sub

Public Function ReadNext() As Boolean
    On Error GoTo HandleError
    
    Dim hFile As Long
    
    If Me.IsOpen Then
        hFile = Me.Handle
        If Not VBA.EOF(hFile) Then
            Line Input #hFile, mstrCurrentLine
            ReadNext = True
        End If
    End If
    
ExitProc:
    Exit Function

HandleError:
    Select Case Err.Number
        Case 62
            ' Just ignore
        Case Else
            ReadNext = False
    End Select
    Resume ExitProc
End Function

Property Get Text() As String
    Text = mstrCurrentLine
End Property

Property Get EOF() As Boolean
    If Me.IsOpen Then
        EOF = VBA.EOF(Me.Handle)
    End If
End Property

Property Get Path() As String
    ' Return the path of the file from the
    ' Private class variable
    Path = mstrPath
End Property

Property Let Path(ByVal strPath As String)
    ' Set the path property of the file.
    ' If a file is already open, close it
    If Me.IsOpen Then
        Me.FileClose
    End If
    mstrPath = strPath
End Property

Property Get Handle() As Long
    ' Get the OS file handle if a file
    ' is open
    If Me.IsOpen Then
        Handle = mhFile
    End If
End Property

Property Get OpenMode() As TextFileOpenMode
    ' Retrieve the open mode of the file
    OpenMode = mlngOpenMode
End Property
Property Let OpenMode(ByVal lngMode As TextFileOpenMode)
    ' Set the open mode of the file only if
    ' a file is not already open
    If Not Me.IsOpen Then
        mlngOpenMode = lngMode
    End If
End Property

Property Get IsOpen() As Boolean
    ' Return value of private variable indicating
    ' whether a fiile is currently open
    IsOpen = mfIsOpen
End Property

Private Sub Class_Initialize()
    ' Set default file open mode property
    Me.OpenMode = tfOpenReadWrite
End Sub

Private Sub Class_Terminate()
    ' If a file is still open then close it
    ' before terminating
    If Me.IsOpen Then
        Me.FileClose
    End If
End Sub

Public Function Exists() As Boolean
    ' Return existance of file based on
    ' a directory search.
    On Error Resume Next
    Exists = (Len(Dir$(Me.Path)) > 0)
    If Err.Number <> 0 Then
        Exists = False
    End If
End Function
