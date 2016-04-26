Attribute VB_Name = "folderbrowse"
Option Compare Database

'************** Code Start **************
'This code was originally written by Terry Kreft.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code courtesy of
'Terry Kreft

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
            "SHGetPathFromIDListA" (ByVal pidl As Long, _
            ByVal pszPath As String) As Long
            
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
            "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
            As Long
            
Private Const BIF_RETURNONLYFSDIRS = &H1


Public Function browsefolder(szDialogTitle As String, buttonindex As Integer) As String
  Dim x As Long, bi As BROWSEINFO, dwIList As Long
  Dim szPath As String, wPos As Integer
  
    With bi
        .hOwner = hWndAccessApp
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(255)
    x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If x Then
        wPos = InStr(szPath, Chr(0))
        browsefolder = Left$(szPath, wPos - 1)
    Else
        browsefolder = vbNullString
    End If
    'MsgBox "length szPath is " & Asc(Left(szPath, 1))
    
    If Not Asc(Left(szPath, 1)) = 0 Then
        'MsgBox "szPath is not null"
        Select Case buttonindex
            Case 0
                If Len(Left(szPath, Len(RTrim(szPath)))) > 80 Then
                    MsgBox "The length of this directory cannot be over 80 characters", vbOKOnly, ""
                Else
                    Forms("frmAdminOptionsCandidates").Controls("txtFileserverPath") = szPath
                End If
            Case 1
                If Len(Left(szPath, Len(RTrim(szPath)))) > 80 Then
                    MsgBox "The length of this directory cannot be over 80 characters", vbOKOnly, ""
                Else
                    Forms("frmMain").Controls("txtClaimLocation") = szPath
                End If
            Case 2
                Forms("frmMain").Controls("txtClaimLocation") = szPath
            Case 3
                Forms("frmMain").Controls("txtClaimLocation") = szPath
        End Select
    End If
            
End Function





Public Function browsefile(strFormName As String, strDir As String, filetype As String)
Dim strFilter As String

Dim cdl As CommonDlg
Set cdl = New CommonDlg

cdl.hwndOwner = Forms(strFormName).hwnd
cdl.CancelError = True

On Error GoTo HandleErrors

cdl.InitDir = strDir
cdl.Filename = ""
cdl.DefaultExt = ""

Select Case filetype
    Case "mdb"
        cdl.Filter = "Access files (*.mdb)|" & "*.mdb,*.mde"
    Case "xls"
        cdl.Filter = "Excel files (*.*)|" & "*.xls;*.xlsx;*.xlsb"
    Case "doc"
        cdl.Filter = "Word files (*.*)|" & "*.doc"
    Case "Custom"
        cdl.Filename = Filename
    Case Else
        cdl.Filter = "All files (*.*)|" & "*.*"
End Select
cdl.FilterIndex = 1
cdl.OpenFlags = cdlOFNEnableHook Or cdlOFNNoChangeDir Or cdlOFNFileMustExist

'cdl.CallBack = adhFnPtrToLong(AddressOf GFNCallback)

cdl.ShowOpen

txtFileOpen = cdl.Filename

''If (cdl.OpenFlags And cdlOFNExtensionDifferent) <> 0 Then
'    MsgBox "You choose a different extension!"
'End If

'MsgBox cdl.FileName
browsefile = cdl.Filename
'Forms("frmEmailMerge").Controls("txtAttachFile") = browsefile


ExitHere:
    Set cdl = Nothing
    Exit Function

HandleErrors:
    Select Case Err.Number
        Case cdlCancel
            Resume ExitHere
        Case Else
            MsgBox "Error: " & Err.Description & _
             "(" & Err.Number & ")"
        End Select
        Resume ExitHere
End Function



