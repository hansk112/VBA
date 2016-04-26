Attribute VB_Name = "ListFile"
Option Compare Database
Option Explicit

'list files to tables
'http://allenbrowne.com/ser-59alt.html

Dim gCount As Long ' added by Crystal

Sub runListFiles()
    'Usage example.
    Dim strPath As String _
    , strFileSpec As String _
    , booIncludeSubfolders As Boolean
    
    strPath = "E:\"
    strFileSpec = "*.*"
    booIncludeSubfolders = True
    
    ListFilesToTable strPath, strFileSpec, booIncludeSubfolders
End Sub

'crystal modified parameter specification for strFileSpec by adding default value
Public Function ListFilesToTable(strPath As String _
    , Optional strFileSpec As String = "*.*" _
    , Optional bIncludeSubfolders As Boolean _
    )
On Error GoTo Err_Handler
    'Purpose:   List the files in the path.
    'Arguments: strPath = the path to search.
    '           strFileSpec = "*.*" unless you specify differently.
    '           bIncludeSubfolders: If True, returns results from subdirectories of strPath as well.
    'Method:    FilDir() adds items to a collection, calling itself recursively for subfolders.
    
    Dim colDirList As New Collection
    Dim varitem As Variant
    Dim rst As DAO.Recordset
    
   Dim mStartTime As Date _
      , mSeconds As Long _
      , mMin As Long _
      , mMsg As String
      
   mStartTime = Now()
   '--------
    
    Call FillDirToTable(colDirList, strPath, strFileSpec, bIncludeSubfolders)
      
   mSeconds = DateDiff("s", mStartTime, Now())
   
   mMin = mSeconds \ 60
   If mMin > 0 Then
      mMsg = mMin & " min "
      mSeconds = mSeconds - (mMin * 60)
   Else
      mMsg = ""
   End If
   
   mMsg = mMsg & mSeconds & " seconds"
   
   'MsgBox "Done adding " & Format(gCount, "#,##0") & " files from " & strPath _
   '   & IIf(Len(Trim(strFileSpec)) > 0, " for file specification --> " & strFileSpec, "") _
   '   & vbCrLf & vbCrLf & mMsg, , "Done"
  
Exit_Handler:
   SysCmd acSysCmdClearStatus
   '--------
    
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, , "ERROR"
    
    'remove next line after debugged -- added by Crystal
    Stop: Resume 'added by Crystal
    Resume Next
    Resume Exit_Handler
End Function

Private Function FillDirToTable(colDirList As Collection _
    , ByVal strFolder As String _
    , strFileSpec As String _
    , bIncludeSubfolders As Boolean)
   
    'Build up a list of files, and then add add to this list, any additional folders
    On Error GoTo Err_Handler
    
    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant
    Dim strSQL As String

    'Add the files to the folder.
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
         gCount = gCount + 1
         SysCmd acSysCmdSetStatus, gCount
         strSQL = "INSERT INTO tblFile " _
          & " (FName, FPath) " _
          & " SELECT """ & strTemp & """" _
          & ", """ & strFolder & """;"
         CurrentDb.Execute strSQL
        colDirList.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Build collection of additional subfolders.
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0& Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop
        'Call function recursively for each subfolder.
        For Each vFolderName In colFolders
            Call FillDirToTable(colDirList, strFolder & TrailingSlash(vFolderName), strFileSpec, True)
        Next vFolderName
    End If

Exit_Handler:
    
    Exit Function

Err_Handler:
    strSQL = "INSERT INTO tblFile " _
    & " (FName, FPath) " _
    & " SELECT ""  ~~~ ERROR ~~~""" _
    & ", """ & strFolder & """;"
    CurrentDb.Execute strSQL
    
    Resume Exit_Handler
End Function

Public Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0& Then
        If Right(varIn, 1&) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function

