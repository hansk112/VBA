Attribute VB_Name = "Module1"
Sub Unhide_Columns()
    'Excel objects.
    Dim m_wbBook As Workbook
    Dim m_wsSheet As Worksheet
    Dim m_rnCheck As Range
    Dim m_rnFind As Range
    Dim m_stAddress As String
    Dim mName As String
    
    'Initialize the Excel objects.
    Set m_wbBook = ThisWorkbook
    Set m_wsSheet = m_wbBook.ActiveSheet
    
    'Search the four columns for any constants.
    Set m_rnCheck = m_wsSheet.Range("A1:Y136").SpecialCells(xlCellTypeConstants)
    
    'Retrieve all columns that contain X. If there is at least one, begin the DO/WHILE loop.
    With m_rnCheck
        Set m_rnFind = .Find(what:="Merchandiser:")
        If Not m_rnFind Is Nothing Then
            m_stAddress = m_rnFind.Address
            Debug.Print m_rnFind.Address
            'Unhide the column, and then find the next X.
            Do
         '       m_rnFind.EntireColumns ("A:D")
                Set m_rnFind = .FindNext(m_rnFind)

                Dim rngFin As String
                rngFin = m_rnFind.Address(ReferenceStyle:=xlA1)
                mName = m_wsSheet.Range(rngFin).Offset(0, 3).Select
                Debug.Print mName
            Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress
        End If
    End With

End Sub


Sub arrTest()


Dim arr() As Variant



End Sub
