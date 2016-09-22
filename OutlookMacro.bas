Attribute VB_Name = "Module2"
Sub searchInbox()

Dim myolapp As New Outlook.Application
Dim myNameSpace As Outlook.NameSpace
Dim myInbox As Outlook.MAPIFolder
Dim myItems As Outlook.Items
Dim objAttachments As Outlook.Attachments
Dim myItem As Object
Dim sn As String
Dim path As String

Set myNameSpace = myolapp.GetNamespace("MAPI")
'set folder to dc folders
Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox).Folders("DC")
Set myItems = myInbox.Items



' http://www.mrexcel.com/forum/general-excel-discussion-other-questions/444884-saving-attachments-using-date-received.html
' loop through inbox
For Each myItem In myItems
 sn = myItem.SenderName
    '   Werner Von Wielligh <Werner.VonWielligh@foodstuffs.co.nz>
    If myItem.Class = olMail Then
        
    ' LNI sender name "Werner Von Wielligh"
        If sn = "Werner Von Wielligh" Then
            Dim LNIDC As String
            LNIDC = myItem.Subject
            saveAttachtoDisk myItem
         '   Found = True
        End If
        
        ' UNI 'weekly sales report string found and email is noreply..
        If InStr(1, myItem.Subject, "Weekly Sales Report") > 0 And sn = "noreply@foodstuffs.co.nz" Then
            Dim UNIDC As String
            UNIDC = myItem.Subject
            saveAttachtoDisk myItem
        End If
        
        ' South Island Goodman Fielder <BW@foodstuffs-si.co.nz>
        Debug.Print sn
        If InStr(1, myItem.Subject, "PNS") > 0 And sn = "Goodman Fielder" Then
            Dim SIPS As String
            SIPS = myItem.Subject
            saveAttachtoDisk myItem
        End If
        
        If InStr(1, myItem.Subject, "NW") > 0 And sn = "Goodman Fielder" Then
            Dim SINW As String
            SINW = myItem.Subject
           saveAttachtoDisk myItem
        End If
        
        If InStr(1, myItem.Subject, "4sq") > 0 And sn = "Goodman Fielder" Then
            Dim SI4Sq As String
            SI4Sq = myItem.Subject
            saveAttachtoDisk myItem
        End If
        
    End If
    
  '  End If
Next myItem

Set myolapp = Nothing
' Call saveAttachtoDisk
msg = MsgBox("Distribution Centre Processing Complete For: " & vbNewLine _
 & UNIDC & path4 & vbNewLine _
 & LNIDC & vbNewLine _
 & SIPS & vbNewLine _
& SINW & vbNewLine _
& SI4Sq & vbNewLine, , _
"DC Files Succesfuly Saved")

End Sub

Private Sub startUpExcel(path As String, fileName As String, sEmail As String, msgSubject As String)

Dim stFol As String
Dim outM As Object
Dim xlapp As Excel.Application
Dim xlWorkbook As Excel.Workbook
Dim sh As Excel.Worksheet

path3 = path

Debug.Print "FileName: " & aName & " " & " email: " & sEmail & " subject Heading:  " & msgSubject
If sEmail = "noreply@foodstuffs.co.nz" Then
fileName = "UNI"
ElseIf sEmail = "Werner.VonWielligh@foodstuffs.co.nz" Then
fileName = "LNI"
ElseIf sEmail = "BW@foodstuffs-si.co.nz" And InStr(1, msgSubject, "PNS") > 0 Then
fileName = "SI PNS"
ElseIf sEmail = "BW@foodstuffs-si.co.nz" And InStr(1, msgSubject, "NW") > 0 Then
fileName = "SI NW"
ElseIf sEmail = "BW@foodstuffs-si.co.nz" And InStr(1, msgSubject, "4sq") > 0 Then
fileName = "SI 4sq"
End If

'' UNZIP SOUTH ISLAND FILE

If path3 = "C:\Users\Hans.Kalders\Desktop\TempDC\ZSD_M01_Q201.ZIP" Then
Unzip1 path3
'GoTo skipToSI:
End If
' new path4 = G:\Fresh\Shared\0_BW_FSDC_SCAN\
'C:\Users\Hans.Kalders\Desktop\TempDC\ZSD_M01_Q201\ZSD_M01_Q201_00000.xls
Debug.Print path3
If fileName = "SI PNS" Or fileName = "SI NW" Or fileName = "SI 4sq" Then
  path3 = "C:\Users\Hans.Kalders\Desktop\TempDC\ZSD_M01_Q201\ZSD_M01_Q201_00000.xls"
  
  path4 = path3
  GoTo skipToSI:
End If
path4 = Left(path3, 22)

skipToSI:
Debug.Print path3
Set xlapp = CreateObject("Excel.Application")
Set xlWorkbook = xlapp.Workbooks.Open(path3)
Dim xlrange As Excel.Range

' do something to spreadhseet
xlapp.Application.Visible = False
With xlapp.ActiveSheet
    If fileName = "UNI" Then
    .Columns("J:J").NumberFormat = "0"
    End If
    If fileName = "LNI" Then
    .Columns("L:M").NumberFormat = "0"
    .Columns("F").NumberFormat = "dd/mm/yyyy"
    ' if format is not equal to then last week data then throw error
    
    End If
    If fileName = "SI PNS" Or fileName = "SI NW" Or fileName = "SI 4sq" Then
    path4 = "C:\Users\Hans.Kalders\"
    
    End If
    
End With

' name as today ddmmyy format name
today = Now() - 1
Debug.Print today
todayString = Left(Format(today, "yymmdd"), 6)
lastweek = Now() - 7
lastweekstring = Left(Format(lastweek, "yymmdd"), 6)
fileConvention = todayString & " - " & lastweekstring & " " & fileName & ".csv"
Debug.Print fileConvention

' save spreadsheet
xlapp.DisplayAlerts = False
xlapp.ActiveWorkbook.SaveAs path4 & fileConvention, xlCSV
xlapp.Application.Quit
xlapp.Quit

End Sub

Function saveAttachtoDisk(itm As Outlook.MailItem)
     Dim objAtt As Outlook.Attachment
     Dim saveFolder As String, dcpath As String, sendEmail As String, asubject As String
     Dim bannerName As String
     saveFolder = "C:\Users\Hans.Kalders\Desktop\TempDC"
     sendEmail = itm.SenderEmailAddress
     asubject = itm.Subject
     For Each objAtt In itm.Attachments

     objAtt.SaveAsFile saveFolder & "\" & objAtt.DisplayName
 '    ElseIf sendEmail = "noreply@foodstuffs.co.nz" Then
 '    bannerName = "UNI"
 '    objAtt.SaveAsFile saveFolder & "\" & objAtt.DisplayName & bannerName
 '    End If
     Dim dName As String
     dName = objAtt.DisplayName
 '    Debug.Print senName
     dcpath = saveFolder & "\" & objAtt.DisplayName
     startUpExcel dcpath, dName, sendEmail, asubject
     
     Set objAtt = Nothing
     Next
End Function
Sub Unzip1(str_FILENAME As Variant)
    Dim oApp As Object
    Dim Fname As Variant
    Dim FnameTrunc As Variant
    Dim FnameLength As Long
 '   Dim xlapp As Excel.Application
 '   Set xlapp = CreateObject("Excel.Application")

 '   str_FILENAME = "C:\Users\Hans.Kalders\Desktop\TempDC\ZSD_M01_Q201.ZIP"
    Fname = str_FILENAME
    FnameLength = Len(Fname)
    FnameTrunc = Left(Fname, FnameLength - 4) & "\"

    If Len(Dir(FnameTrunc, vbDirectory)) = 0 Then
    MkDir FnameTrunc
    End If

    If Fname = False Then
        'Do nothing
    Else
        'Make the new folder in root folder
        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")
  '      xlapp.Application.CutCopyMode = False

        oApp.NameSpace(FnameTrunc).CopyHere oApp.NameSpace(Fname).Items
    End If
    
'    xlapp.Application.DisplayAlerts = False
'    xlapp.Application.Quit
'    xlapp.Quit
    
End Sub
