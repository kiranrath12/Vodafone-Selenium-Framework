'#################################################################################################
' 	Function Name : RetrieveAccount_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveAccount_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sRootProduct = adoData("RootProduct") & ""
	If Ucase(sNVT) = "YES" Then
		sAccount = adoData("Account") & ""
		sMSISDN = adoData("MSISDN") & ""

		If sAccount = "" Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error","Account No is blank"
			Exit Function
		End If
		If sMSISDN = "" Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error","MSISDN is blank"
			Exit Function
		End If

		DictionaryTest_G("AccountNo") = sAccount
		DictionaryTest_G("ACCNTMSISDN") = sMSISDN
		DictionaryTest_G("RootProduct0") = sRootProduct
		DictionaryTempTest_G("RootProduct0") = sRootProduct

	else
		sDataDefinition = adoData("DataDefinition") & ""
		Set vParamsDictionary = CreateObject("Scripting.Dictionary")
		If Ucase(sPool)=Ucase("Synthetic") Then
			vOutput=vDataProxyConnector.TryBlockAndRetrieveAccountDataSynthetic( "15.2", sDataDefinition, vParamsDictionary)
			If (vOutput <> "true") Then
				 iModuleFail = 1
				AddVerificationInfoToResult "Error",vOutput
				Exit Function
			End If
		else
			vOutput=vDataProxyConnector.TryBlockAndRetrieveAccountData( "15.2", sDataDefinition, vParamsDictionary)
			If (vOutput <> "true") Then
				 iModuleFail = 1
				AddVerificationInfoToResult "Error",vOutput
				Exit Function
			End If
		End If
		If DictionaryTest_G.Exists("AccountNo") Then
			DictionaryTest_G.Item("AccountNo")=vParamsDictionary("CUSTOMER_CODE")
		else
			DictionaryTest_G.add "AccountNo",vParamsDictionary("CUSTOMER_CODE")
		End If
	
		If DictionaryTest_G.Exists("ACCNTMSISDN") Then
			DictionaryTest_G.Item("ACCNTMSISDN")=vParamsDictionary("MSISDN")
		else
			DictionaryTest_G.add "ACCNTMSISDN",vParamsDictionary("MSISDN")
		End If
		If vParamsDictionary.Exists("QuerySQL") Then
			AddVerificationInfoToResult "QuerySQL", vParamsDictionary("QuerySQL")
		End If
	End If
	DictionaryTest_G(sRootProduct) = vParamsDictionary("ROOT_PRODUCT")
	DictionaryTempTest_G(sRootProduct) = vParamsDictionary("ROOT_PRODUCT")

	AddVerificationInfoToResult "MSISDN",DictionaryTest_G.Item("ACCNTMSISDN")
	AddVerificationInfoToResult "AccountNo",DictionaryTest_G("AccountNo")
    AddVerificationInfoToResult "RootProduct",DictionaryTest_G(sRootProduct)

	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : RetrieveAccountSecondTime_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveAccountSecondTime_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sDataDefinition = adoData("DataDefinition") & ""
	sRootProduct = adoData("RootProduct") & ""

    Set vParamsDictionary = CreateObject("Scripting.Dictionary")
	If Ucase(sPool)=Ucase("Synthetic") Then
		vOutput=vDataProxyConnector.TryBlockAndRetrieveAccountDataSynthetic( "15.2", sDataDefinition, vParamsDictionary)
		If (vOutput <> "true") Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error",vOutput
			Exit Function
		End If
	else
		vOutput=vDataProxyConnector.TryBlockAndRetrieveAccountData( "15.2", sDataDefinition, vParamsDictionary)
		If (vOutput <> "true") Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error",vOutput
			Exit Function
		End If
	End If
	If DictionaryTest_G.Exists("NewAccountNo") Then
		DictionaryTest_G.Item("NewAccountNo")=vParamsDictionary("CUSTOMER_CODE")
	else
		DictionaryTest_G.add "NewAccountNo",vParamsDictionary("CUSTOMER_CODE")
	End If

	If DictionaryTest_G.Exists("NEWACCNTMSISDN") Then
		DictionaryTest_G.Item("NEWACCNTMSISDN")=vParamsDictionary("MSISDN")
	else
		DictionaryTest_G.add "NEWACCNTMSISDN",vParamsDictionary("MSISDN")
	End If
	DictionaryTest_G("RootProductNew") = vParamsDictionary("ROOT_PRODUCT")


	AddVerificationInfoToResult "NewMSISDN",vParamsDictionary("MSISDN")
	 AddVerificationInfoToResult "RootProduct",DictionaryTest_G("RootProductNew")
	AddVerificationInfoToResult "NewAccountNo",vParamsDictionary("CUSTOMER_CODE")


	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : RetrievePromotionDetails_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrievePromotionDetails_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrievePromotionDetails$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sPromotion = adoData("Promotion") & ""
	sRootProduct = adoData("RootProduct") & ""

	Dim dictProduct
	Set dictProduct = CreateObject("Scripting.Dictionary")
	vOutput=vDataProxyConnector.TryRetrieveProductData("15.2",sPromotion, dictProduct)
	If (vOutput <> "true") Then
'	If (vDataProxyConnector.TryRetrieveProductData("15.2",sPromotion, dictProduct) = False) Then
		 iModuleFail = 1
		AddVerificationInfoToResult "Error",vOutput
		Exit Function

	End If

	If DictionaryTest_G.Exists("PartNo") Then
		DictionaryTest_G.Item("PartNo")=dictProduct("PART_NUMBER")
	else
		DictionaryTest_G.add "PartNo",dictProduct("PART_NUMBER")
	End If

	If DictionaryTest_G.Exists("ProductName") Then
		DictionaryTest_G.Item("ProductName")=dictProduct("PRODUCT_NAME")
	else
		DictionaryTest_G.add "ProductName",dictProduct("PRODUCT_NAME")
	End If
	DictionaryTest_G(sRootProduct) = dictProduct("ROOT_PRODUCT")
	DictionaryTempTest_G(sRootProduct) = dictProduct("ROOT_PRODUCT")

	AddVerificationInfoToResult "ProductName",dictProduct("PRODUCT_NAME")
	AddVerificationInfoToResult "RootProduct",dictProduct("ROOT_PRODUCT")
	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : RetrieveMSISDNAndICCID_fn

' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
'Function RetrieveMSISDNAndICCID_fn1()
'	Dim returnedMSISDN
'	If Not vDataProxyConnector.TryReserveAvailableMSISDN(returnedMSISDN) Then
'		 iModuleFail = 1
'		AddVerificationInfoToResult "Error","MSISDN not available"
'		Exit Function
'	End If
'
'	Dim returnedICCID
'	If Not vDataProxyConnector.TryReserveAvailableICCID(returnedICCID) Then
'		 iModuleFail = 1
'		AddVerificationInfoToResult "Error","ICCID not available"
'		Exit Function
'	End If
'
'
'	If DictionaryTest_G.Exists("MSISDN") Then
'		DictionaryTest_G.Item("MSISDN")=returnedMSISDN
'	else
'		DictionaryTest_G.add "MSISDN",returnedMSISDN
'	End If
'
'	If DictionaryTest_G.Exists("ICCID") Then
'		DictionaryTest_G.Item("ICCID")=returnedICCID
'	else
'		DictionaryTest_G.add "ICCID",returnedICCID
'	End If
'
'
'
'End Function

'#################################################################################################
' 	Function Name : RetrieveRouter_fn

' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Function RetrieveRouter_fn()
	Dim returnedROUTER
	If Not vDataProxyConnector.TryReserveAvailableRouter(returnedROUTER) Then
		 iModuleFail = 1
		AddVerificationInfoToResult "Error","Router numbers not available"
		Exit Function
	End If

	If DictionaryTest_G.Exists("ROUTER") Then
		DictionaryTest_G.Item("ROUTER")=returnedROUTER
	else
		DictionaryTest_G.add "ROUTER",returnedROUTER
	End If

	AddVerificationInfoToResult "Router Number",returnedROUTER

End Function



'#################################################################################################
' 	Function Name : RetrieveMSISDNAndICCID_fn

' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Function RetrieveMSISDNAndICCID_fn()

   If Ucase(sNVT) = "YES" Then
		Call RetrieveMSISDN_fn
	else
		Dim returnedMSISDN
		If Ucase(sPool)=Ucase("Synthetic") Then
			If Not vDataProxyConnector.TryReserveAvailableMsisdnSynthetic(returnedMSISDN) Then
				 iModuleFail = 1
				AddVerificationInfoToResult "Error","MSISDN is not available"
				Exit Function
			End If
		else
			If Not vDataProxyConnector.TryReserveAvailableMSISDN(returnedMSISDN) Then
				 iModuleFail = 1
				AddVerificationInfoToResult "Error","MSISDN  is not available"
				Exit Function
			End If
		End If
		Dim returnedICCID
		If Ucase(sPool)=Ucase("Synthetic") Then
			If Not vDataProxyConnector.TryReserveAvailableIccidSynthetic(returnedICCID) Then
				 iModuleFail = 1
				AddVerificationInfoToResult "Error","ICC id is not available"
				Exit Function
			End If
		else
			If Not vDataProxyConnector.TryReserveAvailableICCID(returnedICCID) Then
				 iModuleFail = 1
				AddVerificationInfoToResult "Error","ICC id is not available"
				Exit Function
			End If
		End If
	
		If DictionaryTest_G.Exists("MSISDN") Then
			DictionaryTest_G.Item("MSISDN")=returnedMSISDN
		else
			DictionaryTest_G.add "MSISDN",returnedMSISDN
		End If
	
		If DictionaryTest_G.Exists("ICCID") Then
			DictionaryTest_G.Item("ICCID")=returnedICCID
		else
			DictionaryTest_G.add "ICCID",returnedICCID
		End If
	
			AddVerificationInfoToResult "MSISDN",returnedMSISDN
			AddVerificationInfoToResult "ICCID",returnedICCID
	End If
End Function


'#################################################################################################
' 	Function Name : RetrieveMSISDNAndICCIDQTP_fn

' 	Description : 
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Function RetrieveMSISDNAndICCIDQTP_fn()

			strConn=ConnectionDetails("QTP")
			set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
			comm.CommandTimeout=120
			conn.Open strConn
			strError = conn.State
		
			If strError <>1 Then		
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , "QTP Database connection not established"  & err.description
				Exit Function
			Else
				AddVerificationInfoToResult  "Info" , "QTP Database connection  established"
			End If


		   If Ucase(sNVT) = "YES" Then
				Call RetrieveMSISDN_fn
			else
				
				Dim returnedMSISDN 
				If Ucase(sPool)=Ucase("Synthetic") Then
					If Not vDataProxyConnector.TryReserveAvailableMsisdnSynthetic(returnedMSISDN) Then
						 iModuleFail = 1
						AddVerificationInfoToResult "Error","MSISDN is not available"
						Exit Function
					End If
				else
						sQuery = "select MSISDN from MSISDN_Repository WHERE STATUS = 'Available' and Environment = '"& sEnv &"' and rownum =1"

						Set recordset = CreateObject("ADODB.Recordset")
						recordset.Open sQuery,conn
						i=0
						Do While Not recordset.EOF			
							returnedMSISDN = recordset.Fields.Item(0)
							i = i+1
							Exit Do
						Loop
					
						 If i=0 Then
								AddVerificationInfoToResult  "Fail" , "No MSISDN available.."
								iModuleFail = 1
								Exit Function
						End If
						On error resume next
						vQuerySQL = "update MSISDN_Repository set status = 'Reserved' where MSISDN ='" & returnedMSISDN & "'"
						Set recordset1 = CreateObject("ADODB.Recordset")
						recordset1.Open vQuerySQL,conn
						recordset.Close
						recordset1.Close
				End If
			'sQuery = "SELECT PART_NUMBER FROM PRODUCT_DATA WHERE ENVIRONMENT='" & sEnv &"' AND KEY='" & sPromotion & "'"


		Dim returnedICCID
			If Ucase(sPool)=Ucase("Synthetic") Then
				If Not vDataProxyConnector.TryReserveAvailableIccidSynthetic(returnedICCID) Then
					 iModuleFail = 1
					AddVerificationInfoToResult "Error","ICC id is not available"
					Exit Function
				End If
		else
						sQuery = "SELECT ICCID from ICCID_Repository WHERE STATUS = 'Available' and Environment = '"& sEnv &"' and rownum =1"

						Set recordset = CreateObject("ADODB.Recordset")
						recordset.Open sQuery,conn
						i=0
						Do While Not recordset.EOF			
							returnedICCID = recordset.Fields.Item(0)
							i = i+1
							Exit Do
						Loop
					
						 If i=0 Then
								AddVerificationInfoToResult  "Fail" , "No ICC ID available.."
								iModuleFail = 1
								Exit Function
						End If
				On error resume next
				vQuerySQL = "update ICCID_Repository set status = 'Reserved' where ICCID ='"& returnedICCID &"'"
				Set recordset1 = CreateObject("ADODB.Recordset")
				recordset1.Open vQuerySQL,conn

				recordset.Close
				recordset1.Close
		
			End If

				err.clear
				conn.Close
			
				If DictionaryTest_G.Exists("MSISDN") Then
					DictionaryTest_G.Item("MSISDN")=returnedMSISDN
				else
					DictionaryTest_G.add "MSISDN",returnedMSISDN
				End If
			
				If DictionaryTest_G.Exists("ICCID") Then
					DictionaryTest_G.Item("ICCID")=returnedICCID
				else
					DictionaryTest_G.add "ICCID",returnedICCID
				End If
			
					AddVerificationInfoToResult "MSISDN",returnedMSISDN
					AddVerificationInfoToResult "ICCID",returnedICCID
			End If
	End Function

'#################################################################################################
' 	Function Name : RetrieveCustomerData_fn

' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Function RetrieveCustomerData_fn(vDictionary1)

'	Dim vDictionary1
'	Set vDictionary1 = CreateObject("Scripting.Dictionary")
'	print vDataProxyConnector.TryRetrieveDynamicData("CUSTOMER-DATA", vDictionary1)
    If (vDataProxyConnector.TryRetrieveDynamicData("CUSTOMER-DATA", vDictionary1) = False) Then
		AddVerificationInfoToResult "Error,""Unable to fetch details."
	End If
	


End Function


'#################################################################################################
' 	Function Name : AccountUnReserveAndReturn_fn

' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Function AccountUnReserveAndReturn_fn()


    If (vDataProxyConnector.TryAccountUnReserveAndReturn(AccountNo, sEnv) = False) Then
		AddVerificationInfoToResult "Info,""Account unreserved failed"
	End If


End Function


'#################################################################################################
' 	Function Name : RetrieveOrderStatus_fn
' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function RetrieveOrderStatus_fn()

	'Get Data
	Dim adoData	  
'    strSQL = "SELECT * FROM [CaptureDataOutputSheet$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\OutputSheet.xls",strSQL)
	sOrderNum = DictionaryTest_G.Item("OrderNo")
'	sEnv = adoData("Env") & ""
'		vStatus=
    	If (vDataProxyConnector.TryGetOrderStatus(sEnv, sOrderNum) = False) Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error","Order not found in database."
			Exit Function
		End If
		sStatus=vDataProxyConnector.TryGetOrderStatus(sEnv, sOrderNum)
		If sStatus = "Complete" Then
			AddVerificationInfoToResult "Info","Order Number :" & sOrderNum & " status is Complete."
		Else
			AddVerificationInfoToResult "Fail","Order Number :" & sOrderNum & " status is not Complete."
			 iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
		End If


		DictionaryTest_G("Status")=sStatus


	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : GetAddress_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function GetAddress_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Address$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sType = adoData("Type") & ""
	sAccountType = adoData("AccountType") & ""


	Set vAddressDictionary = CreateObject("Scripting.Dictionary")


	If (vDataProxyConnector.TryGetAddress( sType, vAddressDictionary, sAccountType) = False) Then
		 iModuleFail = 1
			AddVerificationInfoToResult "Error","Address not found in database."
			Exit Function
	End If


	If DictionaryTest_G.Exists("Address") Then
		DictionaryTest_G.Item("Address")=vAddressDictionary("FULLADDRESS")
	else
		DictionaryTest_G.add "Address",vAddressDictionary("FULLADDRESS")
	End If

	AddVerificationInfoToResult "Address",vAddressDictionary("FULLADDRESS")

	If DictionaryTest_G.Exists("RetnNo") Then
		DictionaryTest_G.Item("RetnNo")=vAddressDictionary("LANDLINE_NMBR")
	else
		DictionaryTest_G.add "RetnNo",vAddressDictionary("LANDLINE_NMBR")
	End If

	AddVerificationInfoToResult "Retention Number",vAddressDictionary("LANDLINE_NMBR")

	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : FreeReservedRouter_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function FreeReservedRouter_fn()

	'Get Data

		If (vDataProxyConnector.TryFreeReservedRouter(sRouter) = False) Then
'			 iModuleFail = 1
			AddVerificationInfoToResult "Info","Router cannot be unreserved."
			Exit Function
		End If

	
'	If Err.Number <> 0 Then
'            iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If
	
End Function


'#################################################################################################
' 	Function Name : FreeReservedMsisdn_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function FreeReservedMsisdn_fn()

	'Get Data

		If (vDataProxyConnector.TryFreeReservedMsisdn(DictionaryTest_G.Item("MSISDN")) = False) Then
'			 iModuleFail = 1
			AddVerificationInfoToResult "Info","MSISDN cannot be unreserved."
			Exit Function
		End If

	
'	If Err.Number <> 0 Then
'            iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If
	
End Function

'#################################################################################################
' 	Function Name : FreeReservedICCID_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function FreeReservedICCID_fn()

	'Get Data

		If (vDataProxyConnector.TryFreeReservedIccid(DictionaryTest_G.Item("ICCID")) = False) Then
'			 iModuleFail = 1
			AddVerificationInfoToResult "Info","ICCID cannot be unreserved."
			Exit Function
		End If

	
'	If Err.Number <> 0 Then
'            iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If
	
End Function

'#################################################################################################
' 	Function Name : RetrievePermanentAccount_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrievePermanentAccount_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sDataDefinition = adoData("DataDefinition") & ""
    Set vParamsDictionary = CreateObject("Scripting.Dictionary")

	If (vDataProxyConnector.TryPermanentBlockAndRetrieveAccountData( "15.2", sDataDefinition, vParamsDictionary) = False) Then
		 iModuleFail = 1
		AddVerificationInfoToResult "Error","Account Data not found"
		Exit Function
	End If

	If DictionaryTest_G.Exists("AccountNo") Then
		DictionaryTest_G.Item("AccountNo")=vParamsDictionary("CUSTOMER_CODE")
	else
		DictionaryTest_G.add "AccountNo",vParamsDictionary("CUSTOMER_CODE")
	End If

	If DictionaryTest_G.Exists("ACCNTMSISDN") Then
		DictionaryTest_G.Item("ACCNTMSISDN")=vParamsDictionary("MSISDN")
	else
		DictionaryTest_G.add "ACCNTMSISDN",vParamsDictionary("MSISDN")
	End If


	AddVerificationInfoToResult "MSISDN",vParamsDictionary("MSISDN")

	AddVerificationInfoToResult "AccountNo",vParamsDictionary("CUSTOMER_CODE")


	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function




'#################################################################################################
' 	Function Name : RetrieveVoucher_fn

' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Function RetrieveVoucher_fn()
	Dim returnedVoucher
	If Not vDataProxyConnector.TryReserveAvailableVoucher(returnedVoucher) Then
		 iModuleFail = 1
		AddVerificationInfoToResult "Error","Voucher numbers not available"
		Exit Function
	End If

	If DictionaryTest_G.Exists("Voucher") Then
		DictionaryTest_G.Item("Voucher")=returnedVoucher
	else
		DictionaryTest_G.add "Voucher",returnedVoucher
	End If

	AddVerificationInfoToResult "Voucher Number",returnedVoucher

End Function


'#################################################################################################
' 	Function Name : Wait_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function Wait_fn()

'		'Get Data
		Dim adoData	  
		strSQL = "SELECT * FROM [Wait$] WHERE  [RowNo]=" & iRowNo
		Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
		sWaitTimeInSec = adoData("WaitTimeInSec") & ""

		   Wait(sWaitTimeInSec)
AddVerificationInfoToResult "WAITED FOR : M ","200 second"
End Function


'#################################################################################################
' 	Function Name : RetrieveTransactData_fn
' 	Description : This Function retrieves Transact data from DB
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function RetrieveTransactData_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveTransactData$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sDataType = adoData("DataType") & ""

	Dim dictProduct
	Set dictProduct = CreateObject("Scripting.Dictionary")
	vOutput=vDataProxyConnector.TryBlockAndRetrieveTransactData(sDataType, dictProduct)
	If (vOutput <> "true") Then
		 iModuleFail = 1
		AddVerificationInfoToResult "Error",vOutput
		Exit Function

	End If

	If DictionaryTest_G.Exists("Title") Then
		DictionaryTest_G.Item("Title")=dictProduct("TITLE")
	else
		DictionaryTest_G.add "Title",dictProduct("TITLE")
	End If

	If DictionaryTest_G.Exists("Firstname") Then
		DictionaryTest_G.Item("Firstname")=dictProduct("FIRSTNAME")
	else
		DictionaryTest_G.add "Firstname",dictProduct("FIRSTNAME")
	End If

	AddVerificationInfoToResult "Firstname",dictProduct("FIRSTNAME")

	If DictionaryTest_G.Exists("Surname") Then
		DictionaryTest_G.Item("Surname")=dictProduct("SURNAME")
	else
		DictionaryTest_G.add "Surname",dictProduct("SURNAME")
	End If

	AddVerificationInfoToResult "Surname",dictProduct("SURNAME")

	If DictionaryTest_G.Exists("DOB") Then
		DictionaryTest_G.Item("DOB")=dictProduct("DOB")
	else
		DictionaryTest_G.add "DOB",dictProduct("DOB")
	End If

	AddVerificationInfoToResult "DOB",dictProduct("DOB")

	If DictionaryTest_G.Exists("Postcode") Then
		DictionaryTest_G.Item("Postcode")=dictProduct("POSTCODE")
	else
		DictionaryTest_G.add "Postcode",dictProduct("POSTCODE")
	End If

	AddVerificationInfoToResult "Postcode",dictProduct("POSTCODE")

	If DictionaryTest_G.Exists("Name") Then
		DictionaryTest_G.Item("Name")=dictProduct("NAME")
	else
		DictionaryTest_G.add "Name",dictProduct("NAME")
	End If

	AddVerificationInfoToResult "Name",dictProduct("NAME")

	If DictionaryTest_G.Exists("Sortcode") Then
		DictionaryTest_G.Item("Sortcode")=dictProduct("SORTCODE")
	else
		DictionaryTest_G.add "Sortcode",dictProduct("SORTCODE")
	End If

	AddVerificationInfoToResult "Sortcode",dictProduct("SORTCODE")

	If DictionaryTest_G.Exists("Accno") Then
		DictionaryTest_G.Item("Accno")=dictProduct("ACCNO")
	else
		DictionaryTest_G.add "Accno",dictProduct("ACCNO")
	End If

	AddVerificationInfoToResult "Accno",dictProduct("ACCNO")
	
	If DictionaryTest_G.Exists("COMP_REG_NUM ") Then
		DictionaryTest_G.Item("COMP_REG_NUM ")=dictProduct("COMP_REG_NUM")
	else
		DictionaryTest_G.add "COMP_REG_NUM ",dictProduct("COMP_REG_NUM")
	End If
	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : RetrieveSTB_fn

' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Function RetrieveSTB_fn()
   Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveSTB$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sKey = adoData("Key") & ""
	
	Dim returnedSTB
	If Not vDataProxyConnector.TryReserveAvailableSTB(returnedSTB) Then
		 iModuleFail = 1
		AddVerificationInfoToResult "Error","STB numbers not available"
		Exit Function
	End If

	If DictionaryTest_G.Exists(sKey) Then
		DictionaryTest_G.Item(sKey)=returnedSTB
	else
		DictionaryTest_G.add sKey,returnedSTB
	End If

	If DictionaryTempTest_G.Exists(sKey) Then
		DictionaryTempTest_G.Item(sKey)=returnedSTB
	else
		DictionaryTempTest_G.add sKey,returnedSTB
	End If

	AddVerificationInfoToResult "STB Number",returnedSTB

End Function


'#################################################################################################
' 	Function Name : RetrieveBRMAccount_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveBRMAccount_fn()
	'--
	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveBRMAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sDataDefinition = adoData("DataDefinition") & ""
	sAccnt = adoData("Accnt") & ""
	sMSISDN = adoData("MSISDN") & ""
	sRootProduct = adoData("RootProduct") & ""
    Set vParamsDictionary = CreateObject("Scripting.Dictionary")
	If Ucase(sPool)=Ucase("Synthetic") Then
		vOutput=vDataProxyConnector.TryBlockAndRetrieveAccountDataSynthetic( "15.2", sDataDefinition, vParamsDictionary)
		If (vOutput <> "true") Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error",vOutput
			Exit Function
		End If
	else
		vOutput=vDataProxyConnector.TryBlockAndRetrieveBRMAccountData( "15.2", sDataDefinition, vParamsDictionary)
'		For loopVar = 1 to 3
'			If Instr(vOutput,"The request channel timed out while waiting") > 0 OR Instr(vOutput,"Object reference not set to an instance") > 0 	Then
'				vOutput=vDataProxyConnector.TryBlockAndRetrieveBRMAccountData( "15.2", sDataDefinition, vParamsDictionary)
'			End If
'
'			If vOutput = "true" Then
'				Exit For
'			End If
'
'		Next

		If (vOutput <> "true") Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error",vOutput
			Exit Function
		End If
	End If

	DictionaryTest_G(sAccnt) = vParamsDictionary("ACCOUNT_NO")
	DictionaryTest_G(sMSISDN) = vParamsDictionary("LOGIN")
	DictionaryTempTest_G("RootProduct0") = sRootProduct

	If sAccnt <> "AccountNo" Then
		DictionaryTempTest_G (sAccnt)=vParamsDictionary("CUSTOMER_CODE")
	elseif sMSISDN <> "ACCNTMSISDN" Then
		DictionaryTempTest_G (sMSISDN)=vParamsDictionary("MSISDN")	
	End If

	DictionaryTest_G(sRootProduct) = vParamsDictionary("ROOT_PRODUCT")
	DictionaryTempTest_G(sRootProduct) = vParamsDictionary("ROOT_PRODUCT")

	AddVerificationInfoToResult "MSISDN",vParamsDictionary("LOGIN")
	AddVerificationInfoToResult "AccountNo",vParamsDictionary("ACCOUNT_NO")
	AddVerificationInfoToResult "RootProduct",DictionaryTest_G(sRootProduct)
	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function
