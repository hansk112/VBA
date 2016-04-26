Attribute VB_Name = "modCustom"
Option Compare Database


Public Sub CprocImportedList()
On Error GoTo HandleErrors


DoCmd.RunSQL "DELETE FROM tblCOLESImportedFiles"

DoCmd.RunSQL "INSERT INTO tblCOLESImportedFiles ( SalesDate, NetSales, FYWK ) " & _
"SELECT tblCOLESBWDataDaily.SaleDate, Sum(tblCOLESBWDataDaily.NetSales) AS SumOfNetSales, tblCOLESBWDataDaily.FYWKGF " & _
"FROM tblCOLESBWDataDaily " & _
"GROUP BY tblCOLESBWDataDaily.SaleDate, tblCOLESBWDataDaily.FYWKGF " & _
"ORDER BY tblCOLESBWDataDaily.SaleDate DESC;"


LprocTempList "", 1, 0
LprocTempList "", 2, 0


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


Public Sub CprocReportShrinkageInitialize()
On Error GoTo HandleErrors

Dim strPCode As String
Dim qdfNew As DAO.QueryDef
Dim strExportFile As String
Dim intMinWk As Long
Dim intMaxWk As Long



strExportFile = Application.CurrentProject.Path & "\export.xls"

LprocTempList "", 1, 0
LprocTempList "", 2, 0
LprocTempList "", 3, 0


LprocDeleteFile strExportFile




' Create a Customer/Product list for the weeks, Customers, Products chosen
' Some Customer/Products are not supplied every week eg some remote stores that keep products in the freezer so
' I need to use 10 weeks of BW data to get all the Customer/Product combinations

' I will assume that if 10 weeks are not chosen, it will use weeks before it that are not chosen, if I choose 3 weeks, it will
' use 7 weeks that are not chosen.
' My assumption is that all the weeks of data are there


DoCmd.RunSQL "DELETE FROM tblReportCustomerProductDate"
DoCmd.RunSQL "DELETE FROM tblReportCustomerProduct"


LprocTempList "SELECT FYWk FROM tblReportCalendar ORDER BY FYWk DESC", 1, 0
If LfnListEmpty(1, 0, 0) = False Then
    intMaxWk = CLng(Forms("frmMain").Controls("lstbox1").Column(0, 0))
    If Forms("frmMain").Controls("lstbox1").ListCount < 10 Then
        If Right(Forms("frmMain").Controls("lstbox1").Column(0, 0), 2) < 10 Then
            
            intMinWk = CLng(Left(Forms("frmMain").Controls("lstbox1").Column(0, 0), 4)) - 1 & CLng(52 - (10 - CInt(Forms("frmMain").Controls("lstbox1").ListCount)) + 1)
        Else
            intMinWk = CLng(Forms("frmMain").Controls("lstbox1").Column(0, 0)) - 10
        End If


        

    Else
        LprocTempList "SELECT FYWk FROM tblReportCalendar ORDER BY FYWk", 1, 0
        intMinWk = CLng(Forms("frmMain").Controls("lstbox1").Column(0, 0))
        
    End If
Else
    MsgBox "Error: There is no week in tblReportCalendar.", vbOKOnly, "Error"
    Exit Sub
End If



' Get all customer/Product combinations for the weeks chosen above
DoCmd.RunSQL "INSERT INTO tblReportCustomerProduct (CustRef, Product, CustRefProduct) " & _
" SELECT DISTINCT tblCOLESCustomers.CustRef, qryCOLESBWDataDailyFYWk.Product, CustRef & " & """" & "/" & """" & " & [Product] AS CustRefProduct " & _
" FROM qryCOLESBWDataDailyFYWk INNER JOIN tblCOLESCustomers ON qryCOLESBWDataDailyFYWk.Customer = tblCOLESCustomers.CustomerNo " & _
" WHERE qryCOLESBWDataDailyFYWk.FYWk>=" & intMinWk & " And qryCOLESBWDataDailyFYWk.FYWk <= " & intMaxWk



' Add the dates chosen to tblReportCustomerProductDate
DoCmd.RunSQL "INSERT INTO tblReportCustomerProductDate (CustRef, Product, SaleDate, CustRefProductDate) " & _
" SELECT DISTINCT tblReportCustomerProduct.CustRef, tblReportCustomerProduct.Product, tblCOLESCalendarFull.StartDate, [tblReportCustomerProduct].[CustRef] & " & """" & "/" & """" & " & [Product] & " & """" & "/" & """" & " & [StartDate] AS CustRefProductDate " & _
" FROM tblReportProducts INNER JOIN (tblReportCustomers INNER JOIN tblReportCustomerProduct ON tblReportCustomers.CustRef = tblReportCustomerProduct.CustRef) ON tblReportProducts.ProductNo = tblReportCustomerProduct.Product, tblCOLESCalendarFull INNER JOIN tblReportCalendar ON (tblCOLESCalendarFull.FY = tblReportCalendar.FY) AND (tblCOLESCalendarFull.Wk = tblReportCalendar.Wk) " & _
" ORDER BY tblReportCustomerProduct.CustRef, tblReportCustomerProduct.Product, tblCOLESCalendarFull.StartDate;"


' now exclude the ones where Avail=0, CO=0, RetDOR=0

' build the result table with CO, Returns, sold out etc, list price


' do db that gets AWS promo, AWS non promo,

' promo, amount discount




'        DoCmd.RunSQL "INSERT INTO tblReportCustomerProductDate (CustRef, Product, SaleDate, SaleDate, CustRefProductDate) " & _
'        " SELECT DISTINCT tblCOLESCustomers.CustRef, qryCOLESBWDataDailyFYWk.Product, qryCOLESBWDataDailyFYWk.SaleDate, [CustRef] & " & """" & "/" & """" & " & [Product] & " & """" & "/" & """" & " & [SaleDate] AS CustRefProductDate, [CustRef] & " & """" & "/" & """" & " & [Product] & " & """" & "/" & """" & " & [SaleDate] AS CustRefProductDate " & _
'        " FROM (qryCOLESBWDataDailyFYWk INNER JOIN tblReportCalendar ON qryCOLESBWDataDailyFYWk.FYWk = tblReportCalendar.FYWk) INNER JOIN tblCOLESCustomers ON qryCOLESBWDataDailyFYWk.Customer = tblCOLESCustomers.CustomerNo " & _
'        " WHERE (((qryCOLESBWDataDailyFYWk.AvailableUnits)<>0) AND ((qryCOLESBWDataDailyFYWk.CarryOver)<>0) AND ((qryCOLESBWDataDailyFYWk.ReturnsDOR)<>0));"



If intCustomerProduct = 1 Or intCustomerProduct = 2 Then
' All customers included
    If intCustomerProduct = 2 Then
    ' some products excluded
        
    Else
    ' All Customers and products
        
    End If
Else
    ' Just use customers chosen
    
End If






' WOW Customers
DoCmd.OpenQuery "qrydboSPWOWCustomersDelete"
DoCmd.OpenQuery "qrydboSPWOWCustomersPopulate"




' BW Data

' Create a pass-through query to delete
DoCmd.OpenQuery "qrydboSPShrinkageBWDataDelete"
DoCmd.OpenQuery "qrydboSPShrinkageBWDataPopulate"
DoCmd.OpenQuery "qrydboSPShrinkageBWDataUpdate"



' Scan data
DoCmd.OpenQuery "qrydboSPWOWShrinkageScanDelete"
DoCmd.OpenQuery "qrydboSPWOWShrinkageScanPopulate"
DoCmd.OpenQuery "qrydboSPWOWShrinkageScanUpdate"





' Exclude or include products
'strProductsInclude = ""
'strProductsExclude = ""


' Choose products to exclude
'DoCmd.RunSQL "DELETE FROM tblProductsExclude"
'DoCmd.OpenForm "frmProductsExclude", WindowMode:=acDialog


'LprocTempList "SELECT ProductNo FROM tblProductsExclude", 1, 0
'intIndex = 0
'While Not IsNull(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) = True And Not Forms("frmMain").Controls("lstbox1").Column(0, intIndex) = ""
'    If strProductsExclude = "" Then
'        strProductsExclude = "Product=" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
'    Else
'        strProductsExclude = strProductsExclude & "  OR Product=" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
'    End If
    
'    intIndex = intIndex + 1
'Wend


' Include products
' Product Exclude and include just sees what string is shortest and uses that
' Include 4 products ie exclude eg 196 products or exclude 4 products include 196 products
'LprocTempList "SELECT ProductNo FROM tblProductsInclude", 1, 0
'intIndex = 0
'While Not IsNull(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) = True And Not Forms("frmMain").Controls("lstbox1").Column(0, intIndex) = ""
'    If strProductsInclude = "" Then
'        strProductsInclude = "Product<>" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
'    Else
'        strProductsInclude = strProductsInclude & " AND Product<>" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
'    End If
    
'    intIndex = intIndex + 1
'Wend


'If Len(strProductsExclude) < Len(strProductsInclude) Then
'    If Len(strProductsExclude) <> 0 Then
'        'LprocQueryPassThrough "WOW Scan", "DELETE FROM dbo.tblWOWShrinkageBWAmounts WHERE " & strProductsExclude
'        DoCmd.OpenQuery "qrydboSPShrinkageBWDataDelete"

'        LprocQueryPassThrough "WOW Scan", "DELETE FROM dbo.tblWOWShrinkageScanData WHERE " & LfnFindAndReplace(strProductsExclude, "Product", "PCode")
'    End If
'Else
'    If Len(strProductsInclude) <> 0 Then
'        'LprocQueryPassThrough "WOW Scan", "DELETE FROM dbo.tblWOWShrinkageBWAmounts WHERE (" & strProductsInclude & ")"
'        DoCmd.OpenQuery "qrydboSPShrinkageBWDataDelete"
        
'        LprocQueryPassThrough "WOW Scan", "DELETE FROM dbo.tblWOWShrinkageScanData WHERE (" & LfnFindAndReplace(strProductsInclude, "Product", "PCode") & ")"
'    End If
'End If




LprocTempList "SELECT TextField FROM dbo_tblWOWSettings WHERE SettingName = " & """" & "ProductsExclude" & """", 1, 0
intIndex = 0
strPCodes = ""
While Not IsNull(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) = True And Not Forms("frmMain").Controls("lstbox1").Column(0, intIndex) = ""
    
    ' create a passthrough query and run this sql
    DoCmd.RunSQL "DELETE FROM tblCOLESShrinkageBWAmounts WHERE Product = " & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
    'DoCmd.RunSQL "DELETE FROM dbo_tblWOWShrinkageBWAmounts WHERE Product = " & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
    
    DoCmd.RunSQL "DELETE FROM tblCOLESShrinkageScanData WHERE PCode = " & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
    'DoCmd.RunSQL "DELETE FROM dbo_tblWOWShrinkageScanData WHERE PCode = " & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex))
    
    intIndex = intIndex + 1
Wend








' create a stored proc for this

' delete products that haven't been chosen
'LprocTempList "SELECT ProductNo FROM dbo_tblWOWShrinkageProduct", 1, 0
'If LfnListEmpty(1, 0, 0) = False Then

'    ' create a passthrough query and run this sql
'    LprocQueryPassThrough "WOW Scan", "DELETE FROM tblWOWShrinkageBWAmounts WHERE tblWOWShrinkageBWAmounts.Product Not In (SELECT ProductNo FROM tblWOWShrinkageProduct)"
    'DoCmd.OpenQuery "qrydboTemp"
'    If strProductsExclude <> "" Then
'        LprocQueryPassThrough "WOW Scan", "DELETE FROM tblWOWShrinkageBWAmounts WHERE tblWOWShrinkageBWAmounts.Product=" & strProductsExclude
'    End If
    'DoCmd.RunSQL "DELETE dbo_tblWOWShrinkageBWAmounts.Product FROM dbo_tblWOWShrinkageBWAmounts WHERE dbo_tblWOWShrinkageBWAmounts.Product Not In (SELECT ProductNo FROM dbo_tblWOWShrinkageProduct);"
    
'    LprocQueryPassThrough "WOW Scan", "DELETE FROM tblWOWShrinkageScanData WHERE tblWOWShrinkageScanData.PCode Not In (SELECT ProductNo FROM tblWOWShrinkageProduct)"
    'DoCmd.OpenQuery "qrydboTemp"
    'DoCmd.RunSQL "DELETE dbo_tblWOWShrinkageScanAmounts.PCode FROM dbo_tblWOWShrinkageScanAmounts WHERE dbo_tblWOWShrinkageScanAmounts.PCode Not In (SELECT ProductNo FROM dbo_tblWOWShrinkageProduct);"
'End If
'LprocTempList "", 1, 0


ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case 70
            MsgBox "You must close " & strExportFile, vbOKOnly, "Close" & strExportFile
            Resume ExitHere
        Case 3010
            MsgBox "You must close " & strExportFile, vbOKOnly, "Close" & strExportFile
            Resume ExitHere
        Case 3265
        ' qry not found
            Resume Next
        Case Else
            MsgBox "Error: " & Err.Description & " (" & Err.Number & ")", vbOKOnly, "Error"
            LprocErrorInsert "CprocReportShrinkageInitialize", "", Err.Description, Err.Number
    End Select
    Resume Next
End Sub


Public Sub CprocDataSource(Weeks As Integer)
' Weeks is the number of weeks we want to do reports for
' This procedure has 2 tables it populates, tblTempCustomerProductDate and tblTempCustomerProduct2
' all the queries use these 2 tables as the data source, so it is a matter of populating these tables
' the reports can be for 4, 6, 12 weeks whatever you want, and it uses the same queries
On Error GoTo HandleErrors

Dim intMaxFYWk As Long
Dim intMinFYWk As Long
Dim dtMaxDate As Date
Dim dtMinDate As Date


' get maximum week
LprocTempList "SELECT MAX(FYWk) FROM tblTempCustomerProduct", 1, 0
If LfnListEmpty(1, 0, 0) = False Then
    intMaxFYWk = CLng(Forms("frmMain").Controls("lstbox1").Column(0, 0))
Else
    intMaxFYWk = 0
    MsgBox "Error: There is something wrong with the data. You must speak to the Database Administrator.", vbOKOnly + vbCritical, "Error"
End If

' get maximum day
LprocTempList "SELECT MAX(SaleDate) FROM tblReportCustomerProductDate", 1, 0
If LfnListEmpty(1, 0, 0) = False Then
    dtMaxDate = CDate(Forms("frmMain").Controls("lstbox1").Column(0, 0))
Else
    dtMaxDate = CDate("01/01/1900")
    MsgBox "Error: There is something wrong with the data. You must speak to the Database Administrator.", vbOKOnly + vbCritical, "Error"
End If

LprocTempList "", 1, 0


' Find min week
If CInt(Right(intMaxFYWk, 2)) < Weeks Then
    intMinFYWk = CLng(CLng(Left(intMaxFYWk, 4)) - 1 & 52 - CInt(Weeks - CInt(Right(intMaxFYWk, 2))) + 1)
Else
    intMinFYWk = intMaxFYWk - Weeks + 1
End If


' Find min date
dtMinDate = dtMaxDate - (Weeks * 7) + 1


DoCmd.RunSQL "SELECT * INTO tblTempCustomerProductDate FROM tblReportCustomerProductDate WHERE SaleDate>=" & LfnMakeUSDate(dtMinDate)

DoCmd.RunSQL "SELECT * INTO tblTempCustomerProduct2 FROM tblTempCustomerProduct WHERE FYWk>=" & intMinFYWk


ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            MsgBox "Error: " & Err.Description & " (" & Err.Number & ")", vbOKOnly, "Error"
    End Select
    'LprocErrorInsert "cmdExport", Me.Name, Err.Description, Err.Number
    Resume Next
End Sub


Public Sub CprocDealLP(txtWeekSD As Date, txtWeekED As Date, strFYWK As String)
On Error GoTo HandleErrors

    LprocTempList "", 1, 0
    LprocTempList "", 2, 0
    LprocTempList "", 3, 0
    
    
    'Insert data from tblDealsTemp to tblDeals based on SQL created below
    
    DoCmd.RunSQL "Delete * from tblCOLESDeal;"
    
    LprocTempList "SELECT * FROM tblProfile1", 1, 0
   
    strSQL = "(Category IS NULL AND SubCategory IS NULL AND Brand IS NULL AND MaterialNo IS NULL)"
    intIndex = 0
    For intIndex = 0 To Forms("frmMain").Controls("lstbox1").ListCount - 1
    'While IsNull(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) = False And Forms("frmMain").Controls("lstbox1").Column(0, intIndex) <> ""

        strSQL = strSQL & " OR (Category = " & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) & " AND SubCategory IS NULL AND Brand IS NULL AND MaterialNo IS NULL)"

        If Not IsNull(Forms("frmMain").Controls("lstbox1").Column(1, intIndex)) And Not Forms("frmMain").Controls("lstbox1").Column(1, intIndex) = "" Then
            strSQL = strSQL & " OR (Category =" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) & " AND SubCategory = " & CLng(Forms("frmMain").Controls("lstbox1").Column(1, intIndex)) & " AND Brand IS NULL AND MaterialNo IS NULL)"
        Else
            If IsNull(Forms("frmMain").Controls("lstbox1").Column(2, intIndex)) Or Forms("frmMain").Controls("lstbox1").Column(2, intIndex) = "" Then
                strSQL = strSQL & " OR (Category =" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) & ")"
            End If
        End If


        If Not IsNull(Forms("frmMain").Controls("lstbox1").Column(2, intIndex)) And Not Forms("frmMain").Controls("lstbox1").Column(2, intIndex) = "" Then
            strSQL = strSQL & " OR (Category =" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) & " AND SubCategory IS NULL AND Brand = " & CLng(Forms("frmMain").Controls("lstbox1").Column(2, intIndex)) & " AND MaterialNo IS NULL)"
        Else
            
        End If
        
        'Debug.Print Forms("frmMain").Controls("lstbox1").Column(0, intIndex) & " " & Forms("frmMain").Controls("lstbox1").Column(1, intIndex) & " " & Forms("frmMain").Controls("lstbox1").Column(2, intIndex)
       
    Next intIndex
 
    
    'Debug.Print "INSERT INTO tblCOLESDeal SELECT tblBAKSDeal.* FROM tblBAKSDeal WHERE " & strsql
    DoCmd.RunSQL "INSERT INTO tblCOLESDeal SELECT tblBAKSDeal.* FROM tblBAKSDeal WHERE " & strSQL

        ' Do subbrands as well if MaterialNo Profile settings is at SubCatNum level
    strSQl2 = ""
    intIndex = 0
    For intIndex = 0 To Forms("frmMain").Controls("lstbox1").ListCount - 1
    'While IsNull(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) = False And Forms("frmMain").Controls("lstbox1").Column(0, intIndex) <> ""
        If IsNull(Forms("frmMain").Controls("lstbox1").Column(2, intIndex)) = True Or Forms("frmMain").Controls("lstbox1").Column(2, intIndex) = "" Then
        
            If IsNull(Forms("frmMain").Controls("lstbox1").Column(1, intIndex)) = False And Forms("frmMain").Controls("lstbox1").Column(1, intIndex) <> "" Then
                LprocTempList "SELECT DISTINCT BrandNum FROM tblBAKSProduct WHERE CategoryNum =" & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) & " AND SubCategoryNum = " & CLng(Forms("frmMain").Controls("lstbox1").Column(2, intIndex)), 2, 0
            Else
                LprocTempList "", 2, 0
            End If
            
            intIndex2 = 0
            While Not IsNull(Forms("frmMain").Controls("lstbox2").Column(0, intIndex2)) = True And Not Forms("frmMain").Controls("lstbox2").Column(0, intIndex2) = ""
                If strSQl2 = "" Then
                    strSQl2 = "(CategoryNum = " & CLng(Forms("frmMain").Controls("lstbox1").Column(0, intIndex)) & " AND BrandNum =" & CLng(Forms("frmMain").Controls("lstbox2").Column(0, intIndex2)) & ")"
                Else
                    strSQl2 = strSQl2 & " OR (CategoryNum = " & CLng(Forms("frmMain").Controls("lstbox1").Column(1, intIndex)) & " AND BrandNum =" & CLng(Forms("frmMain").Controls("lstbox2").Column(0, intIndex2)) & ")"
                End If
                intIndex2 = intIndex2 + 1
            Wend
        End If
    Next intIndex
        'intIndex = intIndex + 1
    'Wend
    
    If strSQl2 <> "" Then
        'Debug.Print "INSERT INTO tblCOLESDeal SELECT tblBAKSDeal.* FROM tblBAKSDeal WHERE " & strSQL2
        DoCmd.RunSQL "INSERT INTO tblCOLESDeal SELECT tblBAKSDeal.* FROM tblBAKSDeal WHERE " & strSQl2
    End If
    
    LprocTempList "", 1, 0
    LprocTempList "", 2, 0
    strSQl2 = ""
    LprocTempList "SELECT DISTINCT Material FROM tblProfile2", 2, 0
    

    ' Do individual sku's as well

    
    For intIndex2 = 0 To Forms("frmMain").Controls("lstbox2").ListCount - 1
        'While Not IsNull(Forms("frmMain").Controls("lstbox2").Column(0, intIndex2)) And Not Forms("frmMain").Controls("lstbox2").Column(0, intIndex2) = ""
            If strSQl2 = "" Then
                strSQl2 = "MaterialNo =" & CLng(Forms("frmMain").Controls("lstbox2").Column(0, intIndex2))
            Else
                strSQl2 = strSQl2 & " OR MaterialNo =" & CLng(Forms("frmMain").Controls("lstbox2").Column(0, intIndex2))
            End If
            Debug.Print CLng(Forms("frmMain").Controls("lstbox2").Column(0, intIndex2))
    Next intIndex2
            'intIndex2 = intIndex2 + 1
        'Wend
        
        
      DoCmd.RunSQL "INSERT INTO tblCOLESDeal SELECT * FROM tblBAKSDeal WHERE " & strSQl2
      
      DoCmd.RunSQL "UPDATE tblCOLESDeal SET [EndDate] = " & LfnMakeUSDate("01/01/2040") & " WHERE Year([EndDate])=9999"
      DoCmd.RunSQL "DELETE FROM tblCOLESDeal WHERE EndDate < " & LfnMakeUSDate(txtWeekED)
      'DoCmd.RunSQL "DELETE FROM tblCOLESDeal WHERE DealEndDate <" & LfnMakeUSDate(txtWeekSD) & " OR DealStartDate >" & LfnMakeUSDate(txtWeekED)
      
'----------------------------------------------------------------------------------------------------------------------------




    ' Now we have all the deals we need
    ' Breakdown the customers. Convert tblDeals -> tblDealsCustomer
    
    DoCmd.RunSQL "DELETE FROM tblCOLESDealCustomer"
    
    DoCmd.OpenQuery "qryDealsCustomerBanner"
    DoCmd.OpenQuery "qryDealsCustomerBannerState"
    DoCmd.OpenQuery "qryDealsCustomerBannerSubSalesRegion"
    DoCmd.OpenQuery "qryDealsCustomerCustomerID"
        
    
    
    DoCmd.RunSQL "DELETE FROM tblCOLESDealCustomerProduct"
    
    DoCmd.OpenQuery "qryDealsCustomerProductCat"
    DoCmd.OpenQuery "qryDealsCustomerProductCatSubCat"
    DoCmd.OpenQuery "qryDealsCustomerProductCatBrand"
    DoCmd.OpenQuery "qryDealsCustomerProductProductNo"

    DoCmd.RunSQL "SELECT DISTINCT tblCOLESDealCustomerProduct.* INTO tblCOLESDealCustomerProduct2 FROM tblCOLESDealCustomerProduct;"
    DoCmd.DeleteObject acTable, "tblCOLESDealCustomerProduct"
    DoCmd.Rename "tblCOLESDealCustomerProduct", acTable, "tblCOLESDealCustomerProduct2"
'----------------------------------------------------------------------------------------------------------------------------
       
        
    DoCmd.RunSQL "DELETE * FROM tblCOLESPriceLPIP"
    
    If strFYWK = "All" Then
     'If this is Calculation of All using tblBWUploadBACKEND
        DoCmd.RunSQL "INSERT INTO tblCOLESPriceLPIP ( Customer, Material ) " & _
        "SELECT tblBWUploadBACKEND.Customer, tblBWUploadBACKEND.Material " & _
        "FROM tblBWUploadBACKEND " & _
        "GROUP BY tblBWUploadBACKEND.Customer, tblBWUploadBACKEND.Material;"
    Else
    'If this calculation of Certain Week using this
         DoCmd.RunSQL "INSERT INTO tblCOLESPriceLPIP ( Customer, Material ) " & _
        "SELECT tblBWUpload.Customer, tblBWUpload.Material " & _
        "FROM tblBWUpload " & _
        "GROUP BY tblBWUpload.Customer, tblBWUpload.Material;"
        
        DoCmd.RunSQL "INSERT INTO tblCOLESPriceLPIP ( Customer, Material ) " & _
        "SELECT DISTINCT tblCOLESBWDataDaily.CustomerNo, tblCOLESBWDataDaily.ProductNo " & _
        "FROM tblCOLESBWDataDaily LEFT JOIN tblCOLESPriceLPIP ON (tblCOLESBWDataDaily.ProductNo = tblCOLESPriceLPIP.Material) AND (tblCOLESBWDataDaily.CustomerNo = tblCOLESPriceLPIP.Customer) " & _
        "WHERE (((tblCOLESPriceLPIP.Customer) Is Null) AND ((tblCOLESPriceLPIP.Material) Is Null) AND ((tblCOLESBWDataDaily.FYWKGF)='" & strFYWK & "'));"

    End If
    
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.CustomerMaterial = tblCOLESPriceLPIP.Customer & '/' & tblCOLESPriceLPIP.Material"
    
    'DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP INNER JOIN tblBAKSCustomer ON tblCOLESPriceLPIP.Customer = tblBAKSCustomer.CustomerNumber SET tblCOLESPriceLPIP.PriceListType = [tblBAKSCustomer].[PriceListType];"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP INNER JOIN tblBAKSCustomer ON tblCOLESPriceLPIP.Customer = tblBAKSCustomer.CustomerNumber SET tblCOLESPriceLPIP.PriceListType = [tblBAKSCustomer].[PriceListType], tblCOLESPriceLPIP.State = [tblBAKSCustomer].[State], tblCOLESPriceLPIP.DeliveryPlant = [tblBAKSCustomer].[DeliveryPlant];"
    
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP INNER JOIN tblBAKSProductPrice ON (tblCOLESPriceLPIP.PriceListType = tblBAKSProductPrice.PriceListType) AND (tblCOLESPriceLPIP.Material = tblBAKSProductPrice.MaterialNo) SET tblCOLESPriceLPIP.LPOrig = [tblBAKSProductPrice].[Price];"
    

    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.LPOrig = 0 WHERE tblCOLESPriceLPIP.LPOrig IS NULL"

    DoCmd.OpenQuery "qryDCustProdDiscZBMP" 'Make table Deal which has ZBMP
    DoCmd.OpenQuery "qryDCustProdBestDealZBMP" 'Find min deal of ZBMP
    'Update tblCOLESDealCustomerProduct
    DoCmd.RunSQL "UPDATE tblDCustProdDiscZBMP INNER JOIN tblCOLESDealCustomerProduct ON (tblDCustProdDiscZBMP.MaterialNo = tblCOLESDealCustomerProduct.MaterialNo) " & _
    "AND (tblDCustProdDiscZBMP.CustomerNo = tblCOLESDealCustomerProduct.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tblDCustProdDiscZBMP].[Rate];"


    DoCmd.OpenQuery "qryDCustProdDiscZCON" 'Make table Deal which has ZCON
    DoCmd.OpenQuery "qryDCustProdBestDealZCON" 'Find min deal of ZCON
    DoCmd.RunSQL "UPDATE tblCOLESDealCustomerProduct INNER JOIN tblDCustProdDiscZCON ON (tblCOLESDealCustomerProduct.MaterialNo = tblDCustProdDiscZCON.MaterialNo) " & _
    "AND (tblCOLESDealCustomerProduct.CustomerNo = tblDCustProdDiscZCON.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tblDCustProdDiscZCON].[Rate];"


    DoCmd.RunSQL "UPDATE tblDCustProdBestDealZCON INNER JOIN tblCOLESPriceLPIP ON tblDCustProdBestDealZCON.CustomerProduct = tblCOLESPriceLPIP.CustomerMaterial SET tblCOLESPriceLPIP.LP = [tblDCustProdBestDealZCON].[MinOfRate];"
    DoCmd.RunSQL "UPDATE tblDCustProdBestDealZBMP INNER JOIN tblCOLESPriceLPIP ON tblDCustProdBestDealZBMP.CustomerProduct = tblCOLESPriceLPIP.CustomerMaterial SET tblCOLESPriceLPIP.IP = [tblDCustProdBestDealZBMP].[MinOfRate];"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP INNER JOIN [qryDCustProdZCON-ZBMP] ON tblCOLESPriceLPIP.CustomerMaterial = [qryDCustProdZCON-ZBMP].CustomerProduct SET tblCOLESPriceLPIP.ZBMP = [qryDCustProdZCON-ZBMP].[ZBMP];"
    
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET IP = LP WHERE LP IS NOT NULL AND IP IS NULL"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET LP = LPOrig WHERE LP IS NULL"


    DoCmd.DeleteObject acTable, "tblDCustProdDiscZBMP"
    DoCmd.DeleteObject acTable, "tblDCustProdDiscZCON"
    DoCmd.DeleteObject acTable, "tblDCustProdBestDealZCON"
    DoCmd.DeleteObject acTable, "tblDCustProdBestDealZBMP"
    
    
    'ZD60 and ZD61
    'DoCmd.OpenQuery "qryDCustProdDiscZD6061"
    
    strSQL = "SELECT tblCOLESDealCustomerProduct.CustomerNo, tblCOLESDealCustomerProduct.MaterialNo, tblCOLESDealCustomerProduct.ConditionType, " & _
    "Format(IIf([ConditionCurrency]='%',[LP]*[Rate]/100,[Rate]),'Standard') AS Discount, tblCOLESDealCustomerProduct.Rate, tblCOLESDealCustomerProduct.ConditionCurrency, " & _
    "[CustomerNo] & '/' & [MaterialNo] AS CustomerProduct, tblCOLESDealCustomerProduct.StartDate, tblCOLESDealCustomerProduct.EndDate INTO tblDCustProdDiscZD6061 " & _
    "FROM tblCOLESDealCustomerProduct INNER JOIN tblCOLESPriceLPIP ON (tblCOLESDealCustomerProduct.CustomerNo = tblCOLESPriceLPIP.Customer) AND (tblCOLESDealCustomerProduct.MaterialNo = tblCOLESPriceLPIP.Material) " & _
    "WHERE (((tblCOLESDealCustomerProduct.ConditionType) = 'ZD60' Or (tblCOLESDealCustomerProduct.ConditionType) = 'ZD61') " & _
    "And ((tblCOLESDealCustomerProduct.StartDate) <=" & LfnMakeUSDate(txtWeekSD) & ") And ((tblCOLESDealCustomerProduct.EndDate) >=" & LfnMakeUSDate(txtWeekSD) & ")) " & _
    "ORDER BY tblCOLESDealCustomerProduct.CustomerNo, tblCOLESDealCustomerProduct.MaterialNo;"
    DoCmd.RunSQL strSQL
    DoCmd.RunSQL "UPDATE tblCOLESDealCustomerProduct INNER JOIN tblDCustProdDiscZD6061 ON (tblDCustProdDiscZD6061.ConditionType = tblCOLESDealCustomerProduct.ConditionType) AND (tblCOLESDealCustomerProduct.MaterialNo = tblDCustProdDiscZD6061.MaterialNo) AND (tblCOLESDealCustomerProduct.CustomerNo = tblDCustProdDiscZD6061.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tblDCustProdDiscZD6061].[Discount];"
    DoCmd.OpenQuery "qryDCustProdBestDealZD6061"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP INNER JOIN tblDCustProdBestDealZD6061 ON tblCOLESPriceLPIP.CustomerMaterial = tblDCustProdBestDealZD6061.CustomerProduct " & _
    "SET tblCOLESPriceLPIP.ZD6061 = Format([tblDCustProdBestDealZD6061].[MaxOfDiscount],'Standard');"
    DoCmd.DeleteObject acTable, "tblDCustProdDiscZD6061"
    DoCmd.DeleteObject acTable, "tblDCustProdBestDealZD6061"
    
    'ZD62
    
    DoCmd.RunSQL "UPDATE tblCOLESDealCustomerProduct INNER JOIN tblCOLESPriceLPIP ON (tblCOLESDealCustomerProduct.CustomerNo = tblCOLESPriceLPIP.Customer) " & _
    "AND (tblCOLESDealCustomerProduct.MaterialNo = tblCOLESPriceLPIP.Material) " & _
    "SET tblCOLESPriceLPIP.ZD62 = IIf([ConditionCurrency]='%',[LP]*[Rate]/100,[Rate]) " & _
    "WHERE (((tblCOLESDealCustomerProduct.ConditionType)='ZD62'));"
    
    
    'ZD64 and ZD65
    'DoCmd.OpenQuery "qryDCustProdDiscZD6465"
    strSQL = "SELECT tblCOLESDealCustomerProduct.CustomerNo, tblCOLESDealCustomerProduct.MaterialNo, tblCOLESDealCustomerProduct.ConditionType, " & _
    "Format(IIf([ConditionCurrency]='%',[LP]*[Rate]/100,[Rate]),'Standard') AS Discount, tblCOLESDealCustomerProduct.Rate, tblCOLESDealCustomerProduct.ConditionCurrency, " & _
    "[CustomerNo] & '/' & [MaterialNo] AS CustomerProduct, tblCOLESDealCustomerProduct.StartDate, tblCOLESDealCustomerProduct.EndDate INTO tblDCustProdDiscZD6465 " & _
    "FROM tblCOLESDealCustomerProduct INNER JOIN tblCOLESPriceLPIP ON (tblCOLESDealCustomerProduct.CustomerNo = tblCOLESPriceLPIP.Customer) AND (tblCOLESDealCustomerProduct.MaterialNo = tblCOLESPriceLPIP.Material) " & _
    "WHERE (((tblCOLESDealCustomerProduct.ConditionType) = 'ZD64' Or (tblCOLESDealCustomerProduct.ConditionType) = 'ZD65') " & _
    "And ((tblCOLESDealCustomerProduct.StartDate) <=" & LfnMakeUSDate(txtWeekSD) & ") And ((tblCOLESDealCustomerProduct.EndDate) >=" & LfnMakeUSDate(txtWeekSD) & ")) " & _
    "ORDER BY tblCOLESDealCustomerProduct.CustomerNo, tblCOLESDealCustomerProduct.MaterialNo;"
    DoCmd.RunSQL strSQL
    DoCmd.RunSQL "UPDATE tblCOLESDealCustomerProduct INNER JOIN tblDCustProdDiscZD6465 ON (tblDCustProdDiscZD6465.ConditionType = tblCOLESDealCustomerProduct.ConditionType) AND (tblCOLESDealCustomerProduct.MaterialNo = tblDCustProdDiscZD6465.MaterialNo) AND (tblCOLESDealCustomerProduct.CustomerNo = tblDCustProdDiscZD6465.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tblDCustProdDiscZD6465].[Discount];"
    DoCmd.OpenQuery "qryDCustProdBestDealZD6465"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP INNER JOIN tblDCustProdBestDealZD6465 ON tblCOLESPriceLPIP.CustomerMaterial = tblDCustProdBestDealZD6465.CustomerProduct " & _
    "SET tblCOLESPriceLPIP.ZD6465 = Format([tblDCustProdBestDealZD6465].[MaxOfDiscount],'Standard');"
    DoCmd.DeleteObject acTable, "tblDCustProdDiscZD6465"
    DoCmd.DeleteObject acTable, "tblDCustProdBestDealZD6465"
    
    'ZD66 and ZD67
    'DoCmd.OpenQuery "qryDCustProdDiscZD6667"
    strSQL = "SELECT tblCOLESDealCustomerProduct.CustomerNo, tblCOLESDealCustomerProduct.MaterialNo, tblCOLESDealCustomerProduct.ConditionType, " & _
    "Format(IIf([ConditionCurrency]='%',[LP]*[Rate]/100,[Rate]),'Standard') AS Discount, tblCOLESDealCustomerProduct.Rate, tblCOLESDealCustomerProduct.ConditionCurrency, " & _
    "[CustomerNo] & '/' & [MaterialNo] AS CustomerProduct, tblCOLESDealCustomerProduct.StartDate, tblCOLESDealCustomerProduct.EndDate INTO tblDCustProdDiscZD6667 " & _
    "FROM tblCOLESDealCustomerProduct INNER JOIN tblCOLESPriceLPIP ON (tblCOLESDealCustomerProduct.CustomerNo = tblCOLESPriceLPIP.Customer) AND (tblCOLESDealCustomerProduct.MaterialNo = tblCOLESPriceLPIP.Material) " & _
    "WHERE (((tblCOLESDealCustomerProduct.ConditionType) = 'ZD66' Or (tblCOLESDealCustomerProduct.ConditionType) = 'ZD67') " & _
    "And ((tblCOLESDealCustomerProduct.StartDate) <=" & LfnMakeUSDate(txtWeekSD) & ") And ((tblCOLESDealCustomerProduct.EndDate) >=" & LfnMakeUSDate(txtWeekSD) & ")) " & _
    "ORDER BY tblCOLESDealCustomerProduct.CustomerNo, tblCOLESDealCustomerProduct.MaterialNo;"
    DoCmd.RunSQL strSQL
    DoCmd.RunSQL "UPDATE tblCOLESDealCustomerProduct INNER JOIN tblDCustProdDiscZD6667 ON (tblDCustProdDiscZD6667.ConditionType = tblCOLESDealCustomerProduct.ConditionType) AND (tblCOLESDealCustomerProduct.MaterialNo = tblDCustProdDiscZD6667.MaterialNo) AND (tblCOLESDealCustomerProduct.CustomerNo = tblDCustProdDiscZD6667.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tblDCustProdDiscZD6667].[Discount];"
    DoCmd.OpenQuery "qryDCustProdBestDealZD6667"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP INNER JOIN tblDCustProdBestDealZD6667 ON tblCOLESPriceLPIP.CustomerMaterial = tblDCustProdBestDealZD6667.CustomerProduct " & _
    "SET tblCOLESPriceLPIP.ZD6667 = Format([tblDCustProdBestDealZD6667].[MaxOfDiscount],'Standard');"
    DoCmd.DeleteObject acTable, "tblDCustProdDiscZD6667"
    DoCmd.DeleteObject acTable, "tblDCustProdBestDealZD6667"
    
    
    
    ' only add Deal ID if IP is null
    
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZD6061 = 0 WHERE ZD6061 is Null"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZD62 = 0 WHERE ZD62 is Null"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZD6465 = 0 WHERE ZD6465 is Null"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZD6667 = 0 WHERE ZD6667 is Null"
    
    
    
    ' Update ZBMP Discount
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZBMP = Format([LP]-ABS([ZD6061])-ABS([ZD62])-ABS([ZD6465])-ABS([ZD6667]),'Standard') " & _
    "WHERE (((tblCOLESPriceLPIP.LP) Is Not Null) AND ((tblCOLESPriceLPIP.IP) Is Not Null) AND ((tblCOLESPriceLPIP.LP)<>[IP]));"
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZBMP = 0 WHERE tblCOLESPriceLPIP.ZBMP Is Null;"
    
    ' Work out invoice price
    ' this is needed before deferred deals are calculated, Method can be PI
    
    DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET IP = LP-ABS([ZD6061])-ABS([ZD62])-ABS([ZD6465])-ABS([ZD6667]) WHERE IP IS NULL"


'-------------------------------------------------------------------------------------------------------------------------------

With CurrentDb
        .QueryDefs.Delete ("qryDealsCustomerProduct")
        Set qdfNew = .CreateQueryDef("qryDealsCustomerProduct", "SELECT * FROM tblCOLESDealCustomerProduct WHERE StartDate<=" & LfnMakeUSDate(txtWeekSD) & " AND EndDate>=" & LfnMakeUSDate(txtWeekSD))
End With


' Non National
Dim RebateNo As Integer
For RebateNo = 60 To 69
    
    
    
    strSQL = "SELECT qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery, Sum(IIf([ConditionCurrency]='%',[IP]*[Rate]/100,[Rate])) AS Discount INTO tbl" & RebateNo & _
    " FROM tblCOLESPriceLPIP INNER JOIN qryDealsCustomerProduct ON (tblCOLESPriceLPIP.Material = qryDealsCustomerProduct.MaterialNo) AND (tblCOLESPriceLPIP.Customer = qryDealsCustomerProduct.CustomerNo) " & _
    "GROUP BY qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery " & _
    "HAVING (((qryDealsCustomerProduct.ConditionType)='ZR" & RebateNo & "') AND ((qryDealsCustomerProduct.Bakery) Is NOT Null));"

    
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESPriceLPIP ON (tbl" & RebateNo & ".MaterialNo = tblCOLESPriceLPIP.Material) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESPriceLPIP.Customer) " & _
    "SET tblCOLESPriceLPIP.ZR" & RebateNo & " = [tbl" & RebateNo & "].[Discount];"
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESDealCustomerProduct ON (tbl" & RebateNo & ".ConditionType = tblCOLESDealCustomerProduct.ConditionType) AND (tbl" & RebateNo & ".MaterialNo = tblCOLESDealCustomerProduct.MaterialNo) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESDealCustomerProduct.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tbl" & RebateNo & "].[Discount];"
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZR" & RebateNo & " = 0 WHERE ZR" & RebateNo & " is Null"
    DoCmd.RunSQL strSQL
    
    DoCmd.DeleteObject acTable, "tbl" & RebateNo
    
Next RebateNo

'National
For RebateNo = 60 To 69
    
    
     strSQL = "SELECT qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery, Sum(IIf([ConditionCurrency]='%',[IP]*[Rate]/100,[Rate])) AS Discount INTO tbl" & RebateNo & _
    " FROM tblCOLESPriceLPIP INNER JOIN qryDealsCustomerProduct ON (tblCOLESPriceLPIP.Material = qryDealsCustomerProduct.MaterialNo) AND (tblCOLESPriceLPIP.Customer = qryDealsCustomerProduct.CustomerNo) " & _
    " GROUP BY qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery " & _
    " HAVING (((qryDealsCustomerProduct.ConditionType)='ZR" & RebateNo & "') AND ((qryDealsCustomerProduct.Bakery) Is Null));"

    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESPriceLPIP ON (tbl" & RebateNo & ".MaterialNo = tblCOLESPriceLPIP.Material) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESPriceLPIP.Customer) " & _
    "SET tblCOLESPriceLPIP.[ZR" & RebateNo & "-Nat] = [tbl" & RebateNo & "].[Discount];"
    DoCmd.RunSQL strSQL
  
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESDealCustomerProduct ON (tbl" & RebateNo & ".ConditionType = tblCOLESDealCustomerProduct.ConditionType) AND (tbl" & RebateNo & ".MaterialNo = tblCOLESDealCustomerProduct.MaterialNo) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESDealCustomerProduct.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tbl" & RebateNo & "].[Discount] + NZ(tblCOLESDealCustomerProduct.Discount,0);"
    DoCmd.RunSQL strSQL
    
  
    strSQL = "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.[ZR" & RebateNo & "-Nat] = 0 WHERE [ZR" & RebateNo & "-Nat] is Null"
    DoCmd.RunSQL strSQL
    
    DoCmd.DeleteObject acTable, "tbl" & RebateNo
Next RebateNo

'ZR6A & ZR6B Non National
For RebateNo = 65 To 66
    
    
    
    
     strSQL = "SELECT qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery, Sum(IIf([ConditionCurrency]='%',[IP]*[Rate]/100,[Rate])) AS Discount INTO tbl" & RebateNo & _
    " FROM tblCOLESPriceLPIP INNER JOIN qryDealsCustomerProduct ON (tblCOLESPriceLPIP.Material = qryDealsCustomerProduct.MaterialNo) AND (tblCOLESPriceLPIP.Customer = qryDealsCustomerProduct.CustomerNo) " & _
    " GROUP BY qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery " & _
    " HAVING (((qryDealsCustomerProduct.ConditionType)='ZR6" & Chr(RebateNo) & "') AND ((qryDealsCustomerProduct.Bakery) Is NOT Null));"
    
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESPriceLPIP ON (tbl" & RebateNo & ".MaterialNo = tblCOLESPriceLPIP.Material) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESPriceLPIP.Customer) " & _
    "SET tblCOLESPriceLPIP.ZR6" & Chr(RebateNo) & " = [tbl" & RebateNo & "].[Discount];"
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESDealCustomerProduct ON (tbl" & RebateNo & ".ConditionType = tblCOLESDealCustomerProduct.ConditionType) AND (tbl" & RebateNo & ".MaterialNo = tblCOLESDealCustomerProduct.MaterialNo) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESDealCustomerProduct.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tbl" & RebateNo & "].[Discount];"
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.ZR6" & Chr(RebateNo) & " = 0 WHERE ZR6" & Chr(RebateNo) & " is Null"
    DoCmd.RunSQL strSQL
    
    DoCmd.DeleteObject acTable, "tbl" & RebateNo
Next RebateNo

'ZR6A & ZR6B National
For RebateNo = 65 To 66
    
    
    strSQL = "SELECT qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery, Sum(IIf([ConditionCurrency]='%',[IP]*[Rate]/100,[Rate])) AS Discount INTO tbl" & RebateNo & _
    " FROM tblCOLESPriceLPIP INNER JOIN qryDealsCustomerProduct ON (tblCOLESPriceLPIP.Material = qryDealsCustomerProduct.MaterialNo) AND (tblCOLESPriceLPIP.Customer = qryDealsCustomerProduct.CustomerNo) " & _
    " GROUP BY qryDealsCustomerProduct.CustomerNo, qryDealsCustomerProduct.MaterialNo, qryDealsCustomerProduct.ConditionType, qryDealsCustomerProduct.Bakery " & _
    " HAVING (((qryDealsCustomerProduct.ConditionType)='ZR6" & Chr(RebateNo) & "') AND ((qryDealsCustomerProduct.Bakery) Is Null));"
    
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESPriceLPIP ON (tbl" & RebateNo & ".MaterialNo = tblCOLESPriceLPIP.Material) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESPriceLPIP.Customer) " & _
    "SET tblCOLESPriceLPIP.[ZR6" & Chr(RebateNo) & "-Nat] = [tbl" & RebateNo & "].[Discount];"
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tbl" & RebateNo & " INNER JOIN tblCOLESDealCustomerProduct ON (tbl" & RebateNo & ".ConditionType = tblCOLESDealCustomerProduct.ConditionType) AND (tbl" & RebateNo & ".MaterialNo = tblCOLESDealCustomerProduct.MaterialNo) AND (tbl" & RebateNo & ".CustomerNo = tblCOLESDealCustomerProduct.CustomerNo) SET tblCOLESDealCustomerProduct.Discount = [tbl" & RebateNo & "].[Discount] + NZ(tblCOLESDealCustomerProduct.Discount,0);"
    DoCmd.RunSQL strSQL
    
    strSQL = "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.[ZR6" & Chr(RebateNo) & "-Nat] = 0 WHERE [ZR6" & Chr(RebateNo) & "-Nat] is Null"
    DoCmd.RunSQL strSQL
    
    DoCmd.DeleteObject acTable, "tbl" & RebateNo

Next RebateNo



' Calculate NASP
DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.TDPrice = Round([IP]-ABS([ZR60])-ABS([ZR60-Nat])-ABS([ZR61])-ABS([ZR61-Nat]),2);"
DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.NASP = Format([IP]-ABS([ZR60])-ABS([ZR60-Nat])-ABS([ZR61])-ABS([ZR61-Nat])-ABS([ZR62])-ABS([ZR62-Nat])-ABS([ZR63])-ABS([ZR63-Nat])-ABS([ZR64])-ABS([ZR64-Nat])-ABS([ZR65])-ABS([ZR65-Nat])-ABS([ZR66])-ABS([ZR66-Nat])-ABS([ZR67])-ABS([ZR67-Nat])-ABS([ZR68])-ABS([ZR68-Nat])-ABS([ZR69])-ABS([ZR69-Nat])-ABS([ZR6A])-ABS([ZR6A-Nat])-ABS([ZR6B])-ABS([ZR6B-Nat]),'Standard');"
DoCmd.RunSQL "UPDATE tblCOLESPriceLPIP SET tblCOLESPriceLPIP.NASPOriginal = IIf([LPOrig] = 0, 0,Format(([NASP]*100)/[LPOrig]," & """" & "Standard" & """" & "))"







ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            MsgBox "Error: " & Err.Description & " (" & Err.Number & ")", vbOKOnly, "Error"
    End Select
    'LprocErrorInsert "cmdExport", Me.Name, Err.Description, Err.Number
    Resume Next


End Sub

