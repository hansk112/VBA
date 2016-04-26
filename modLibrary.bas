Attribute VB_Name = "modLibrary"
Option Compare Database

Public intGroupBy As Byte
Public intCustomerProduct As Byte


Public Function LfnIsFormLoaded(formname As String) As Boolean
On Error Resume Next
'Dim fIsOpen As Boolean
LfnIsFormLoaded = CurrentProject.AllForms(formname).IsLoaded
End Function


Function LfnMakeUSDate(x As Variant)
      If Not IsDate(x) Then Exit Function
      LfnMakeUSDate = "#" & Month(x) & "/" & Day(x) & "/" & Year(x) & "#"
End Function

Function LfnMakeUSDt(x As Variant)
      If Not IsDate(x) Then Exit Function
      LfnMakeUSDt = Month(x) & "/" & Day(x) & "/" & Year(x)
End Function

Public Function cleanStr(ByVal dirtyString As String) As String
Dim oRegex As Object

    If oRegex Is Nothing Then Set oRegex = CreateObject("vbscript.regexp")
    With oRegex
        .Global = True
        'Allow A-Z, a-z, 0-9, a space and a hyphen -
        .Pattern = "[^A-Za-z0-9 -]"
        cleanStr = .Replace(dirtyString, vbNullString)
    End With
    cleanStr = WorksheetFunction.Trim(cleanStr)
End Function

Public Sub LprocLotusNotesEmail(emailaddress As String, EmailSubject As String, EmailBody As String, MyAttachment As String)
On Error GoTo HandleErrors

Dim session, db, NotesAttach, NotesDoc As Object
Dim server, mailfile As String
Dim objAttach, objAttach2 As Object


Set session = CreateObject("Notes.NotesSession")
    
server = session.GetEnvironmentString("MailServer", True)
mailfile = session.GetEnvironmentString("MailFile", True)

Set db = session.GetDatabase(server, mailfile)


Set NotesDoc = db.createdocument


Call NotesDoc.replaceItemValue("Subject", EmailSubject)
Call NotesDoc.replaceItemValue("SendTo", emailaddress)
Call NotesDoc.replaceItemValue("Form", "memo")
Call NotesDoc.replaceItemValue("Body", EmailBody)


Set objAttach = NotesDoc.CreateRichTextItem("Attachment")

If Not MyAttachment = "" Then
    Set objAttach2 = objAttach.embedobject(1454, "", MyAttachment)
End If


    
NotesDoc.SAVEMESSAGEONSEND = True
NotesDoc.Send False


Set session = Nothing  'close connection to free memory
Set db = Nothing
Set NotesAttach = Nothing
Set NotesDoc = Nothing
Set objNotesField = Nothing
Set objAttach = Nothing
Set objAttach2 = Nothing
    
   
ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case 7000
            Resume Next
        Case Else
            MsgBox "Error: " & _
             Err.Description & _
             " (" & Err.Number & ")"
    End Select
    'LprocErrorInsert "LprocLotusNotesEmail", "", Err.Description, Err.Number
    Resume Next
End Sub


Public Sub LprocTempList(ByRef rs As String, listnumber As Integer, tabletype As Byte)
'tabletype - True=Table/Query,1-Value List,2-Field List
On Error GoTo HandleErrors

'intTempList = intTempList + 1


Select Case tabletype
    Case 0
        Forms("frmMain").Controls("lstbox" & listnumber).RowSourceType = "Table/Query"
    Case 1
        Forms("frmMain").Controls("lstbox" & listnumber).RowSourceType = "Value List"
    Case 2
        Forms("frmMain").Controls("lstbox" & listnumber).RowSourceType = "Field List"
End Select

Forms("frmMain").Controls("lstbox" & listnumber).RowSource = rs


ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            MsgBox "Error: " & _
             Err.Description & _
             " (" & Err.Number & ")"
    End Select
    Resume ExitHere
    Resume Next
End Sub


Public Function LfnWeekStartDate(dategiven As Date)
Select Case Weekday(dategiven)
    Case 1
    ' Sunday
        LfnWeekStartDate = dategiven - 6
    Case 2
    ' Monday
        LfnWeekStartDate = dategiven
    Case 3
    ' Tuesday
        LfnWeekStartDate = dategiven - 1
    Case 4
    'Wednesday
        LfnWeekStartDate = dategiven - 2
    Case 5
    ' Thursday
        LfnWeekStartDate = dategiven - 3
    Case 6
    ' Friday
        LfnWeekStartDate = dategiven - 4
    Case 7
    ' Saturday
        LfnWeekStartDate = dategiven - 5
End Select

End Function

Public Sub LprocErrorInsert(ByRef procname As String, ByRef formname As String, ByRef errdescription As String, ByRef errnumber As Double)
On Error GoTo HandleErrors

'DoCmd.SetWarnings False

Dim rst3 As ADODB.Recordset

Set rst3 = New ADODB.Recordset

rst3.ActiveConnection = CurrentProject.Connection

rst3.Open "SELECT BoolField1 FROM tblLocalOptions"

If rst3!BoolField1 = True Then
' in debug mode, this is will act as a breakpoint
    If Right(Application.CurrentProject.Name, 3) = "mdb" Then
        MsgBox "error, put a breakpoint here"
    End If
End If

rst3.Close


' insert the error
DoCmd.RunSQL "INSERT INTO tblCOLESError (controlname,formname,errordescription,errornumber,ErrorDate,ErrorTime) VALUES (" & """" & procname & """" & "," & """" & formname & """" & "," & """" & FindAndReplace(errdescription, """", "'") & """" & "," & errnumber & "," & LfnMakeUSDate(Date) & "," & """" & Time & """" & ")"


ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case 2004
            ' not enough memory to perform this task
            Resume Next
        Case Else
            MsgBox "Error: " & _
            Err.Description & _
             " (" & Err.Number & ")"
            'DoCmd.RunSQL "INSERT INTO tblError (controlname,formname,errordescription,errornumber,ErrorDate,ErrorTime) VALUES (" & """" & " " & """" & "," & """" & "LprocErrorInsert" & """" & "," & """" & FindAndReplace(Err.Description, """", "'") & """" & "," & Err.Number & "," & LfnMakeUSDate(Date) & "," & """" & Time & """" & ")"
            'Debug.Print Err.Number & " " & Err.Description
    End Select
    Resume Next
End Sub


Public Function LfnListEmpty(listnumber As Integer, x As Long, y As Long)

'If Forms("frmMain").Controls("lstbox" & listnumber).ListCount = 0 Then
If IsNull(Forms("frmMain").Controls("lstbox" & listnumber).Column(x, y)) = False And Forms("frmMain").Controls("lstbox" & listnumber).Column(x, y) <> "" Then
    LfnListEmpty = False
Else
    LfnListEmpty = True
End If

End Function


Public Function LfnWeekdayName(dt As Date)
Select Case Weekday(dt)
    Case 1
        LfnWeekdayName = "Sunday"
    Case 2
        LfnWeekdayName = "Monday"
    Case 3
        LfnWeekdayName = "Tuesday"
    Case 4
        LfnWeekdayName = "Wednesday"
    Case 5
        LfnWeekdayName = "Thursday"
    Case 6
        LfnWeekdayName = "Friday"
    Case 7
        LfnWeekdayName = "Saturday"
End Select

End Function


Function FindAndReplace(ByVal strInString As String, _
        strFindString As String, _
        strReplaceString As String) As String
Dim intPtr As Integer
    If Len(strFindString) > 0 Then  'catch if try to find empty string
        Do
            intPtr = InStr(strInString, strFindString)
            If intPtr > 0 Then
                FindAndReplace = FindAndReplace & Left(strInString, intPtr - 1) & _
                                        strReplaceString
                    strInString = Mid(strInString, intPtr + Len(strFindString))
            End If
        Loop While intPtr > 0
    End If
    FindAndReplace = FindAndReplace & strInString
End Function


Public Function LfnFilesFromFolder(pFolderPath As String, pFileType As String, pSubDirs As Boolean) As Variant
' This lists the files in a folder
' option to include a filetype eg *.mdb or *.xls
' option to list subdirectories


'Dim A() As Variant
'Dim B() As Variant

' To show results from here
'A = LfnFilesFromFolder("C:\ProjectDB\", "*.mdb", False)
' if you want to sort by name
'B = LfnSortFileNameOrder(A)
    
'Select Case IsArray(Y)
'    Case True 'files found
'        MsgBox UBound(B)
'        For i = LBound(B) To UBound(B)
'            MsgBox B(i)
'        Next i
'    Case False 'no files found
'        MsgBox "No matching files"
'End Select


On Error GoTo HandleErrors


Dim FileArray() As Variant
Dim FileCount As Integer
Dim Filename As String

Dim strFolderSubDir As String
Dim strDirFile As String
Dim strFileName As String
Dim colFolderList As New Collection
Dim intFileCount As Long

intFileCount = 0
strFileList = ""


' Add backslash at end of path if not present
pFolderPath = Trim(pFolderPath)
If Right(pFolderPath, 1) <> "\" Then
    pFolderPath = pFolderPath & "\"
End If


strFileName = Dir(pFolderPath & pFileType)
If strFileName = "" Then
    LfnFilesFromFolder = False
Else
    Do While strFileName <> ""
        intFileCount = intFileCount + 1
        ReDim Preserve FileArray(1 To intFileCount)
        FileArray(intFileCount) = strFileName
        strFileName = Dir()
    Loop


'    If pSubDirs = True Then

'        strFolderSubDir = Dir(pFolderPath & "*", vbDirectory)
'        Do While strFolderSubDir <> ""
'            If strFolderSubDir <> "." And strFolderSubDir <> ".." Then
'                If (GetAttr(pFolderPath & strFolderSubDir) And vbDirectory) = 16 Then
                    'MsgBox pFolderPath & strFolderSubDir
'                    colFolderList.Add pFolderPath & strFolderSubDir & "\"
'                End If
'            End If

'            strFolderSubDir = Dir
'        Loop
        
'    End If
    
    
'    For Each CollectionItem In colFolderList
'        If strFileList = "" Then
'            strFileList = LfnFilesFromFolder(CStr(CollectionItem), pFileType, pSubDirs)
'        Else
'            strFileList = strFileList & ";" & LfnFilesFromFolder(CStr(CollectionItem), pFileType, pSubDirs)
            'Debug.Print strFileList
'        End If
'    Next

'    LfnFilesFromFolder = strFileList


    LfnFilesFromFolder = FileArray
    Exit Function

End If



    ' To get the results:
'    X = LfnFilesFromFolder(p)
'    Select Case IsArray(X)
'        Case True 'files found
'            MsgBox UBound(X)
'            Sheets("Sheet1").Range("A:A").Clear
'            For i = LBound(X) To UBound(X)
'                MsgBox X(i)
'            Next i
'        Case False 'no files found
'            MsgBox "No matching files"
'    End Select


ExitHere:
    Exit Function

HandleErrors:
    LfnFilesFromFolder = False
    LprocErrorInsert "LfnFilesFromFolder", "", Err.Description, Err.Number
    Resume ExitHere
End Function


Public Sub LprocDOSCommand(DOSString As String)
Call Shell(Environ$("COMSPEC") & " /c " & DOSString, vbNormalFocus)
'Environ$("COMSPEC") returns the path to Command.com on the machine and the "/c" argument makes sure that the Dos window is automatically closed when the batch file finishes executing.
End Sub


Public Sub LprocPause(ByVal pSng_Secs As Single)
'Wait for the number of seconds given by pSng_Secs
Dim lSng_Start As Single
Dim lSng_End As Single
On Error GoTo Err_Pause
    lSng_Start = Timer
    lSng_End = Timer + pSng_Secs
    Do While Timer < lSng_End
    '' Correction if the timer moves over to a new day (midnight)
    '' 86400-num of secs in a day
        If Timer < lSng_Start Then lSng_End = lSng_End - 86400
    Loop
Err_Pause:
    Exit Sub
End Sub


Public Sub LprocWriteToFile(strLine As String, fName As String)
' We need modules Line, Lines, TextFile, TextFile2 and TextFile3
On Error GoTo HandleErrors


Dim colLines As Collection
Dim mobjFileTo As TextFile3

Set mobjFileTo = New TextFile3
Set colLines = New Collection
                
mobjFileTo.Path = fName
mobjFileTo.OpenMode = TextFileOpenMode3.tfOpenReadWrite
        
For cLines = 1 To mobjFileTo.Lines.Count
    mobjFileTo.Lines.Remove cLines
Next
        
mobjFileTo.Lines.Add strLine
        
mobjFileTo.FileSave
mobjFileTo.FileClose

Set mobjFileTo = Nothing
Set colLines = Nothing


ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            MsgBox "Error: " & Err.Description & _
             "(" & Err.Number & ")"
            'LprocErrorInsert "Form_Open", Me.Name, Err.Description, Err.Number
            Resume Next
            Resume ExitHere
    End Select
End Sub


Function LfnFindAndReplace(ByVal strInString As String, _
        strFindString As String, _
        strReplaceString As String) As String
Dim intPtr As Integer
    If Len(strFindString) > 0 Then  'catch if try to find empty string
        Do
            intPtr = InStr(strInString, strFindString)
            If intPtr > 0 Then
                LfnFindAndReplace = LfnFindAndReplace & Left(strInString, intPtr - 1) & _
                                        strReplaceString
                    strInString = Mid(strInString, intPtr + Len(strFindString))
            End If
        Loop While intPtr > 0
    End If
    LfnFindAndReplace = LfnFindAndReplace & strInString
End Function


Public Function LprocDeleteFile(filespec As String)
   Set fso = CreateObject("Scripting.FileSystemObject")
   fso.DeleteFile (filespec)
End Function



Public Sub LprocEnumerate_Table(strInput As String)
On Error GoTo ERROR_PROC
Dim aryFields()
Dim lngCount As Long
    
    lngCount = 0
    'strInput = InputBox("Please enter the name of the table for which" & vbCrLf & _
    '        "you wish to list FieldNames & Descriptions." & vbCrLf & vbCrLf & _
    '        "Output will be placed in tab-delimited text file.", "Table Name Input", "MainTabl")
    
    
    If StrPtr(strInput) = 0 Or Len(strInput) = 0 Then
        Exit Sub
    Else
        strSQL = "SELECT * FROM " & strInput
        Dim Adofl As ADODB.Field
        Dim rs As New ADODB.Recordset
        rs.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        For Each Adofl In rs.Fields
            lngCount = lngCount + 1
            ReDim Preserve aryFields(2, lngCount)
            aryFields(1, lngCount) = Adofl.Name
            aryFields(2, lngCount) = LfnGetFieldDesc_ADO(strInput, Adofl.Name)
        Next
        rs.Close 'recordset closed for next item.
    End If ' Feed array to designated file
    If lngCount > 0 Then
        aFileNum = FreeFile
        Open "C:\Temp\TableStruc.txt" For Output As #aFileNum
        For i = 1 To UBound(aryFields(), 2)
        Print #aFileNum, aryFields(1, i) & vbTab & aryFields(2, i)
        Next
        Close #aFileNum
Final_Results:
        Btn = MsgBox("Table fieldnames with Descriptions stored in:" & vbCrLf & vbCrLf & _
            "C:\Temp\TableStruc.txt" & vbCrLf & vbCrLf & _
            "Do you wish to OPEN using NOTEPAD?", vbOKCancel + vbQuestion, _
            " Table Enumeration")
        If Btn = vbOK Then ' Opening the file in Notepad
            Shell "Notepad.exe" & " " & "C:\Temp\TableStruc.txt", vbMaximizedFocus
        End If
    Else
        MsgBox "Table Not Found!", vbOKOnly + vbExclamation, _
            "Bad Table Name"
    End If
    Exit Sub
ERROR_PROC:
    rs.Close
    MsgBox "Error encountered attempting to enumerate table!"
End Sub

' This function requires a reference to ADO 2.5 Ext. for DDL & Security (or higher)

Public Function LfnGetFieldDesc_ADO(ByVal MyTableName As String, ByVal MyFieldName As String)

On Error GoTo Err_GetFieldDescription
 
 
Dim MyDB As New ADOX.Catalog
Dim MyTable As ADOX.Table
Dim MyField As ADOX.Column
   
    MyDB.ActiveConnection = CurrentProject.Connection
    Set MyTable = MyDB.Tables(MyTableName)
    GetFieldDesc_ADO = MyTable.Columns(MyFieldName).Properties("Description")
    Set MyDB = Nothing

Bye_GetFieldDescription:
    Exit Function
Err_GetFieldDescription:
    Beep
    MsgBox Err.Description, vbExclamation
    GetFieldDescription = Null
    Resume Bye_GetFieldDescription
End Function

Public Function fnFileLocked(strFileName As String) As Boolean
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      'MsgBox "Error #" & Str(Err.Number) & " - " & Err.Description
      fnFileLocked = True
      Err.Clear
   End If
End Function

