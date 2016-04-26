VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFile2"
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
Private mlngOpenMode As TextFileOpenMode2
Private mfIsOpen As Boolean

' Enumeration for file open mode
Enum TextFileOpenMode2
    tfOpenReadOnly
    tfOpenReadWrite
    tfOpenAppend
    tfOpenReadOnlyExists
End Enum

' Class object variables
Public Lines As Lines

Public Function FileOpen()
    On Error GoTo HandleError
    
    Dim strLine As String
    
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
    
    ' Read all lines into the Lines collection
    Set Lines = New Lines
    If LOF(mhFile) > 0 Then
        Do Until EOF(mhFile)
            Line Input #mhFile, strLine
            Me.Lines.Add strLine
        Loop
    End If
    
    ' Reset the changed property of all lines
    Me.Lines.Changed = False
    
ExitProc:
    Exit Function

HandleError:
    Select Case Err.Number
        Case 62 ' Input passed EOF
                ' Just ignore this
        Case Else
            FileOpen = False
    End Select
    Resume ExitProc
End Function

Public Sub FileClose()
    ' If we have a valid handle then close the file
    If Me.Handle <> 0 Then
        Close Me.Handle
        mfIsOpen = False
    End If
End Sub

Property Get Path() As String
    ' Return the path of the file from the
    ' Private class variable
    Path = mstrPath
End Property

Property Let Path(ByVal strPath As String)
    ' Set the path property of the file--
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

Property Get OpenMode() As TextFileOpenMode2
    ' Retrieve the open mode of the file
    OpenMode = mlngOpenMode
End Property
Property Let OpenMode(ByVal lngMode As TextFileOpenMode2)
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

Public Function FileSave(Optional ByVal varNewPath As Variant) As Boolean
    Dim strPath As String
    Dim strRename As String
    Dim hFile As Long
    Dim cLine As Long
    Dim objFile As TextFile
    
    ' Make sure file is open in read/write mode
    If Me.IsOpen And Me.OpenMode = tfOpenReadWrite Then
        
        ' Close the file
        Me.FileClose
        
        ' Get the path to save the file to
        If IsMissing(varNewPath) Then
            strPath = Me.Path
        Else
            strPath = varNewPath
        End If
        
        ' Compute name to rename file to temporarily
        strRename = Left(strPath, Len(strPath) - 1) & "~"
        Set objFile = New TextFile
        objFile.Path = strRename
        If objFile.Exists Then
            Kill strRename
        End If
        
        ' Rename existing file
        Name Me.Path As strRename
        
        ' Write Lines collection to new file
        hFile = FreeFile
        Open strPath For Output Access Write As hFile
        For cLine = 1 To Me.Lines.Count
            Print #hFile, Me.Lines.Item(cLine).Text
        Next
        Close hFile
        
        ' Write was successful do delete original
        If objFile.Exists Then
            Kill strRename
        End If
    
        ' Re-open the file
        Me.FileOpen
        
        FileSave = True
    End If
End Function

Public Function Exists() As Boolean
    ' Return existance of file based on
    ' a directory search
    On Error Resume Next
    Exists = (Len(Dir$(Me.Path)) > 0)
    If Err.Number <> 0 Then
        Exists = False
    End If
End Function
