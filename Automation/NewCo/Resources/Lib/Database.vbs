
Public Function ConnectionDetails(Env)

   Select Case Env

'   Case "Jenkin"
'
'		strHost = "10.78.195.156"
'		strPort = "3306"
'		strServiceName = "vdd"
'		strUserName = "root"
'		strPassword = "Admin@123"
'		dbName="vdd"
'
'        strConnection = "DRIVER={MySQL ODBC 5.1 Driver}; Server=" & strHost & "; DATABASE="& dbName& ";uid=root; pwd=Admin@123"
'
''		strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
''												")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
''		strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
''														  ";User ID=" & strUserName & ";Password=" & strPassword & ";"


   Case "E7"

		strHost = "10.78.195.74"
		strPort = "1522"
		strServiceName = "DEVCRM"
		strUserName = "SIEBEL"
		strPassword = "SIEBEL"
		strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
												")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
		strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
														  ";User ID=" & strUserName & ";Password=" & strPassword & ";"

	Case "E4"
			strHost = "10.78.193.202"
			strServiceName = "DEVCRM"
			strPort = "1522"
			strUserName = "SIEBEL"
			strPassword = "SIEBEL"
			strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
													")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
			strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
															  ";User ID=" & strUserName & ";Password=" & strPassword & ";"

	Case "C2"
			strHost = "10.78.221.7"
			strPort = "1522"
			strServiceName = "DEVCRM"
			strUserName = "SIEBEL"
			strPassword = "SIEBEL"
			strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
								")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
			strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
										  ";User ID=" & strUserName & ";Password=" & strPassword & ";"


		Case "E6"
			strHost = "10.78.195.202"
			strPort = "1521"
			strServiceName = "DEVCRM"
			strUserName = "SIEBEL"
			strPassword = "SIEBEL"
			strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
								")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
			strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
										  ";User ID=" & strUserName & ";Password=" & strPassword & ";"


				Case "E2"
			strHost = "10.78.193.12"
			strPort = "1522"
			strServiceName = "DEVCRM"
			strUserName = "SIEBEL"
			strPassword = "SIEBEL"
			strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
								")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
			strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
										  ";User ID=" & strUserName & ";Password=" & strPassword & ";"


		Case "SUP02"
			strHost = "10.78.196.94"
			strPort = "1522"
			strServiceName = "DEVCRM"
			strUserName = "SIEBEL"
			strPassword = "SIEBEL"
			strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
								")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
			strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
										  ";User ID=" & strUserName & ";Password=" & strPassword & ";"
										  
		Case "SUP01"
			strHost = "10.78.194.7"
			strPort = "1522"
			strServiceName = "DEVCRM"
			strUserName = "SIEBEL"
			strPassword = "SIEBEL"
			strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
								")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
			strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
										  ";User ID=" & strUserName & ";Password=" & strPassword & ";"								  

		Case "BRME7"

		strHost = "NewVoE7-dbbrm01"
		strPort = "1521"
		strServiceName = "DEVBRM"
		strUserName = "pin"
		strPassword = "pin"
		strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
												")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
		strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
														  ";User ID=" & strUserName & ";Password=" & strPassword & ";"

	   Case "BRME4"

		strHost = "NewVoE4-dbbrm01"
		strPort = "1521"
		strServiceName = "DEVBRM"
		strUserName = "pin"
		strPassword = "pin"
		strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
												")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
		strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
														  ";User ID=" & strUserName & ";Password=" & strPassword & ";"


		  Case "BRMC2"

		strHost = "newcoc2-dbbrm01"
		strPort = "1521"
		strServiceName = "DEVBRM"
		strUserName = "pin"
		strPassword = "pin"
		strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
												")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
		strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
														  ";User ID=" & strUserName & ";Password=" & strPassword & ";"

		Case "QTP"
			strHost = "qtp.cqrpvr4944ct.eu-west-1.rds.amazonaws.com"
			strPort = "1521"
			strServiceName = "QTP"
			strUserName = "datamanagement"
			strPassword = "D2tamanagement"
			strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
								")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
			strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
										  ";User ID=" & strUserName & ";Password=" & strPassword & ";"

	Case "BRMSUP02"

		strHost = "dbbrm01"
		strPort = "1521"
		strServiceName = "DEVBRM"
		strUserName = "pin"
		strPassword = "pin"
		strDBDesc = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST="& strHost & _
												")(PORT=" &strPort &"))(CONNECT_DATA= (SID=" & strServiceName&")))"
		strConn = "Provider=OraOLEDB.Oracle;Data Source=" & strDBDesc & _
														  ";User ID=" & strUserName & ";Password=" & strPassword & ";"

   End Select
	ConnectionDetails=strConn

End Function
		


'#################################################################################################
' 	Function Name : ExecuteDBQuery_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ExecuteDBQuery_fn()

	strSQL = "SELECT * FROM [ExecuteDBQuery$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Database.xls",strSQL)
	sDB = adoData( "DB")  & ""
	strConn=ConnectionDetails(sDB & sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	conn.Open strConn
	strError = conn.State
		
	If strError <>1 Then

		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , sDB & sEnv & "Database connection not established" & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , sDB & sEnv & "Database connection  established"
	End If

	Do while Not adoData.Eof
'		DictionaryTest_G("AccountNo") = "7000240728"
''		DictionaryTest_G("CDRSeq") = 2370816
'		DictionaryTest_G("ACCNTMSISDN") = "447387966852"
'		DictionaryTest_G("MSISDN") = "447387944193"
'		DictionaryTest_G("RowId") = "1-500HEOO"

		sDD = adoData( "DD")  & ""
		
		sResult = adoData( "Result")  & ""
		If DictionaryTest_G.Exists(sResult) Then
			sResult = Replace (sResult,sResult,DictionaryTest_G(sResult))
		elseif DictionaryTempTest_G.Exists(sResult) Then
			sResult = Replace (sResult,sResult,DictionaryTempTest_G(sResult))
		End If
		If isNumeric(sResult) Then
			sResult =Cdbl(Round(sResult,6))
		End If

		If sResult = "Sysday" Then
			sResult = Day(Now)
		End If
		DictionaryTest_G ("DBResult") = sResult

		sKeys = adoData( "Keys")  & ""
		sKeysArr = Split (sKeys,";")
		sAction = adoData( "Action")  & ""
        sSecondPart = adoData( "SecondPart") & ""


		strSQL1 = "SELECT * FROM [Queries$] WHERE  [DD]='" & sDD & "'"
		Set adoData1 = ExecExcQry(sProductDir & "Data\Database.xls",strSQL1)

		sQuery = adoData1( "Query")  & ""
		For i=0 to Ubound(sKeysArr)
			If DictionaryTest_G.Exists(sKeysArr(i)) Then
				sQuery = Replace (sQuery,"Key"&i+1,DictionaryTest_G (sKeysArr(i)))
				AddVerificationInfoToResult  sKeysArr (i),DictionaryTest_G (sKeysArr(i))

			elseif DictionaryTempTest_G.Exists(sKeysArr(i)) Then
				sQuery = Replace (sQuery,"Key"&i+1,DictionaryTempTest_G (sKeysArr(i)))
				AddVerificationInfoToResult  sKeysArr (i),DictionaryTempTest_G (sKeysArr(i))

			else
				sQuery = Replace (sQuery,"Key"&i+1,sKeysArr(i))
				
			End If
	
		Next
		AddVerificationInfoToResult  "Query",sQuery
		'sEnv = adoData( "Env")  & ""
		
		Set recordset = CreateObject("ADODB.Recordset")
		recordset.Open sQuery,conn
'		sResultArr = Split(sResult,"+")
'		Dim sActResultArr
'		ReDim sActResultArr(recordset.RecordCount)
		i=0	
		If sAction = "Capture" Then

			Do While Not recordset.EOF
						
				For each field in recordset.Fields
	'				If Cstr(sResultArr(i)) = Cstr(field.Value) Then
	'					AddVerificationInfoToResult  "Pass" , "Value is " & sResult & " as expected"
	'				else
	'					AddVerificationInfoToResult  "Fail" , "Actual value is " & field.Name & " but expected is " & sResult
	'				End If
					If field.Name="BDOM" or sSecondPart="Y" Then
						DictionaryTempTest_G(field.Name & i)=field.Value
					End If
					DictionaryTest_G(field.Name & i)=field.Value
					AddVerificationInfoToResult  "Info" , field.Name & i & " - " & DictionaryTest_G(field.Name & i)
				Next
				i=i+1
				recordset.MoveNext
			Loop
			If i=0 Then
				AddVerificationInfoToResult  "Fail" , "No rows returned"
				iModuleFail = 1
			End If
		elseIf sAction = "Add" Then
'			sActResult = 0
'			Dim sActResult
			Do While Not recordset.EOF
						
				For each field in recordset.Fields
					If isEmpty(sActResult) Then
						sActResult = 0
					End If
'                    f=Cdbl(Replace(Cstr(field.Value),".",""))
				
'					sActResult = sActResult + f
					sActResult = Round(sActResult,6) + Round(Cdbl(field.Value),6)
					sActResult = Round(sActResult,6)
					AddVerificationInfoToResult  "Info" , "Row value is " & Round(Cdbl(field.Value),6)
				Next

				recordset.MoveNext
			Loop
			If isNumeric(sResult) = False Then
				AddVerificationInfoToResult  "Fail" , "Actual value is " &sActResult & " but expected is " & sResult
				iModuleFail = 1
			End If
			If Cstr(Cdbl(sActResult)) = Cstr(Cdbl(sResult)) Then
				AddVerificationInfoToResult  "Pass" , "Value is " & sResult & " as expected"
			else
				AddVerificationInfoToResult  "Fail" , "Actual value is " &sActResult & " but expected is " & sResult
				iModuleFail = 1
			End If

		elseIf sAction = "Compare" Then
			sResultArr = Split(sResult,"+")

			Do While Not recordset.EOF
			
			
				For each field in recordset.Fields
					If instr(sResultArr(i),"-")=1 Then
						sResultArr(i) = Replace (sResultArr(i),"-","")
						If DictionaryTest_G.Exists(sResultArr(i)) Then
							sResultArr(i) = Replace (sResultArr(i),sResultArr(i),DictionaryTest_G(sResultArr(i)))
						End If
						sResultArr(i) = "-" & sResultArr(i) 
					End If
					If DictionaryTest_G.Exists(sResultArr(i)) Then
						sResultArr(i) = Replace (sResultArr(i),sResultArr(i),DictionaryTest_G(sResultArr(i)))
					End If

					If Cstr(sResultArr(i)) = Cstr(field.Value) Then
						AddVerificationInfoToResult  "Pass" , "Value is " & sResultArr(i) & " as expected"
					else
						AddVerificationInfoToResult  "Fail" , "Actual value is " & field.Value & " but expected is " & sResultArr(i)
						iModuleFail = 1
					End If
'					DictionaryTest_G(field.Name & i)=field.Value
				Next
				i=i+1
				recordset.MoveNext
			Loop
			If i <> Ubound(sResultArr) + 1 Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , "Expected is " & Ubound(sResultArr)+1 & " rows but actual is " & i
			End If
		End If
		On error resume next
		recordset.Close
		err.clear
		If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
		End If
		adoData.movenext
	Loop
	conn.Close


End Function



'#################################################################################################
' 	Function Name : ExecuteDBQueryOrderCancel_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ExecuteDBQueryOrderCancel_fn()



'	strSQL = "SELECT * FROM [ExecuteDBQuery$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Database.xls",strSQL)
'
'	sQuery = adoData( "Query")  & ""
'	'sEnv = adoData( "Env")  & ""
'	sRecordsAffected = adoData( "RecordsAffected")  & ""
'	If (instr(1,sQuery,"RowId",1)>=1) Then
'		sQuery=Replace(sQuery,"RowId","'" & DictionaryTest_G.Item("RowId") & "'")
'	End If
'	if (instr(1,sQuery,"OrderNo",1)>=1) Then
'		sQuery=Replace(sQuery,"OrderNo","'" & DictionaryTest_G.Item("OrderNo") & "'")
'	End If
'	if (instr(1,sQuery,"MSISDN",1)>=1) Then
'		sQuery=Replace(sQuery,"MSISDN","'" & DictionaryTest_G.Item("MSISDN") & "'")
'	End If

	orderNo=DictionaryTest_G.Item("OrderNo")
	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	conn.Open strConn
	strError = conn.State
		
	If strError <>1 Then

		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Database connection not established"
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Database connection  established"
	End If

    sQuery="UPDATE siebel.S_ORDER_ITEM SET STATUS_CD = 'Complete' WHERE order_id in (select so.row_id from s_order so where so.order_num='"&orderNo&"')"
	Set recordset = conn.Execute (sQuery, RecordsAffected)

    sQuery="UPDATE siebel.S_ORDER SET STATUS_CD = 'Complete' WHERE row_id =(select so.row_id from s_order so where  so.order_num='"&orderNo&"')"
	Set recordset = conn.Execute (sQuery, RecordsAffected)
    sQuery="UPDATE S_ORDER_ITEM SET FULFLMNT_STATUS_CD = 'Complete' WHERE ORDER_ID in (select so.row_id from s_order so where  so.order_num='"&orderNo&"')"
	Set recordset = conn.Execute (sQuery, RecordsAffected)
'	conn.Execute(sQuery)
	Set recordset = conn.Execute (sQuery, RecordsAffected)
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

'	If Cint(sRecordsAffected) >=RecordsAffected and RecordsAffected <> 0 Then
		AddVerificationInfoToResult  "Pass" , RecordsAffected & " row(s) updated"
'	else
'		iModuleFail = 1
'		AddVerificationInfoToResult  "Error" , "Expected rows updated: " & sRecordsAffected &   " Actual rows updated: " & RecordsAffected
'	End If

End Function

Public Function GetPendingandOpenOrdersOnAccount()


	strSQL = "SELECT * FROM [Account$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Database.xls",strSQL)

	sAccount = adoData( "Account")  & ""
	sMSISDN = adoData( "MSISDN")  & ""

	If DictionaryTest_G.Exists("AccountNo") Then
		DictionaryTest_G.Item("AccountNo")=sAccount
	else
		DictionaryTest_G.add "AccountNo",sAccount
	End If

	If DictionaryTest_G.Exists("ACCNTMSISDN") Then
		DictionaryTest_G.Item("ACCNTMSISDN")=sMSISDN
	else
		DictionaryTest_G.add "ACCNTMSISDN",sMSISDN
	End If

   	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	conn.Open strConn
	strError = conn.State

	If DictionaryTest_G.Exists("AccountNo") Then
		DictionaryTest_G.Item("AccountNo")=sAccount
	else
		DictionaryTest_G.add "AccountNo",sAccount
	End If
		
	If strError <>1 Then

		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Database connection not established"
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Database connection  established"
	End If

	sQuery="select so.order_num from s_org_ext soe,s_order so where soe.X_VF_CUSTOMER_CODE in ('"&sAccount&"') and so.ACCNT_ID = soe.row_id and so.status_cd in ('Pending','Open')"
'	sQuery="select so.order_num,so.status_cd from s_org_ext soe,s_order so where soe.X_VF_CUSTOMER_CODE in (select soe.X_VF_CUSTOMER_CODE from s_order_item soi,s_order so, s_org_ext soe where service_num='441968494558'  and soi.order_id = so.row_id and so.ACCNT_ID = soe.row_id) and so.ACCNT_ID = soe.row_id and so.status_cd in ('Pending','Open')"
	Set recordset = conn.Execute (sQuery)

	Do While not recordset.EOF
		orderNo=recordset.Fields.Item(0)
		sQuery1="UPDATE siebel.S_ORDER_ITEM SET STATUS_CD = 'Cancelled' WHERE order_id in (select so.row_id from s_order so where so.order_num='"&orderNo&"')"
        conn.Execute (sQuery1)

		sQuery1="UPDATE siebel.S_ORDER SET STATUS_CD = 'Cancelled' WHERE row_id =(select so.row_id from s_order so where  so.order_num='"&orderNo&"')"
		conn.Execute (sQuery1)
		sQuery1="UPDATE S_ORDER_ITEM SET FULFLMNT_STATUS_CD = 'Cancelled' WHERE ORDER_ID in (select so.row_id from s_order so where  so.order_num='"&orderNo&"')"
		conn.Execute (sQuery1)
		recordset.MoveNext

	Loop
	conn.Close
	Set conn=Nothing

End Function



'#################################################################################################
' 	Function Name : RetrieveAccountQTP_fn()
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveAccountQTP_fn1()

'	strSQL = "SELECT * FROM [ExecuteDBQuery$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Database.xls",strSQL)
'	sDB = adoData( "DB")  & ""
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sDataDefinition = adoData("DataDefinition") & ""
	sAccnt = adoData("Accnt") & ""
	sMSISDN = adoData("MSISDN") & ""
	sRootProduct = adoData("RootProduct") & ""

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
	sQuery = "SELECT QUERY_SQL, QUERY_TYPE, PROMOTION_KEY FROM MINING_QUERY WHERE ENVIRONMENT='" & sEnv &"'  AND QUERY_NAME='" & sDataDefinition & "'"
    Set recordset = CreateObject("ADODB.Recordset")
	recordset.Open sQuery,conn
	i=0
	Do While Not recordset.EOF			
		vQuerySQL = recordset.Fields.Item(0)
		vQueryType = recordset.Fields.Item(1)
		vPromotionKey = recordset.Fields.Item(2)
		i = i+1
		Exit Do
	Loop

	 If i=0 Then
		AddVerificationInfoToResult  "Fail" , "QUERY_NAME not found"
		iModuleFail = 1
	End If
	On error resume next
	recordset.Close
	err.clear
	conn.Close

	If vPromotionKey<>"" Then
		Dim dictProduct
		Set dictProduct = CreateObject("Scripting.Dictionary")
		vOutput=vDataProxyConnector.TryRetrieveProductData("15.2",vPromotionKey, dictProduct)
		If (vOutput <> "true") Then
	'	If (vDataProxyConnector.TryRetrieveProductData("15.2",sPromotion, dictProduct) = False) Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error",vOutput & " Promotion  key not retrieved"
			Exit Function
	
		End If
		DictionaryTest_G("ProductName")= dictProduct("PRODUCT_NAME")
		AddVerificationInfoToResult "Promotion Name",dictProduct("PRODUCT_NAME")
		If dictProduct.Exists ("ROOT_PRODUCT") Then
			DictionaryTest_G(sRootProduct) = dictProduct("ROOT_PRODUCT")
			DictionaryTempTest_G(sRootProduct) = dictProduct("ROOT_PRODUCT")
			vQuerySQL = Replace(vQuerySQL,"RootProduct",dictProduct ("ROOT_PRODUCT"))
		End If
		vQuerySQL = vQuerySQL + " and exists (select 1 from s_asset sa, s_prod_int spi where spi.name = '" + dictProduct("PRODUCT_NAME") + "' and spi.row_id = sa.prod_id and sa.owner_accnt_id = soe.row_id and sa.status_cd = 'Active' and sip.row_id = sa.bill_profile_id) AND ( ( soe.url NOT LIKE '#B') OR ( soe.url IS NULL ) OR ( soe.url = 'Corrupt'))"
'		vQuerySQL = vQuerySQL + "AND sa2.serial_num NOT BETWEEN '447000000000'AND '447099999999'"
		vQuerySQL = Replace(vQuerySQL,"SELECT soe.row_id", "SELECT soe.x_vf_customer_code CUSTOMER_CODE")
		 
	End If

    vQuerySQL = vQuerySQL + " and sa2.created > sysdate-25 and rownum = 1 order by soe.created desc "
	AddVerificationInfoToResult "Query",vQuerySQL

	strConn=ConnectionDetails(sEnv)
	comm.CommandTimeout=120
'	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	conn.Open strConn
	strError = conn.State
		
	If strError <>1 Then

		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established" & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If
	recordset.Open vQuerySQL,conn
	i=0
	Do While Not recordset.EOF
		DictionaryTest_G("AccountNo") = recordset.Fields.Item(0)
'		DictionaryTest_G("ACCNTMSISDN") = recordset.Fields.Item(1)
		i=i+1
		Exit Do
    Loop
	On error resume next
	recordset.Close
	err.clear
	If i=0 Then
		AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
		iModuleFail = 1
		conn.Close
		Exit Function
	End If

'	vQuerySQL = "SELECT soe.x_vf_customer_code, sa2.SERIAL_NUM from s_org_ext soe, s_asset sa2 where sa2.owner_accnt_id = soe.row_id and sa2.SERIAL_NUM like ('44%') and soe.row_id ='" &rowId& "'"
	vQuery = "select distinct z.service_num from s_org_ext soe , s_order o, s_order_item z where o.customer_id = soe.row_id and o.row_id = z.order_id and soe.x_vf_customer_code in ('"+DictionaryTest_G("AccountNo")+"') and z.service_num like '4%'"
	recordset.Open vQuery,conn
	i=0
	Do While Not recordset.EOF
'		DictionaryTest_G("AccountNo") = recordset.Fields.Item(0)
		DictionaryTest_G("ACCNTMSISDN") = recordset.Fields.Item(0)
		i=i+1
		recordset.MoveNext
'		Exit Do
    Loop

	On error resume next
	recordset.Close
	err.clear
'	conn.Close

    If i=0 Then
		AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
		iModuleFail = 1
		conn.Close
		Exit Function
	End If
	vQuerySQL = "update s_org_ext soe set soe.url='#B' where soe.x_vf_customer_code ='"+DictionaryTest_G("AccountNo")+"'"
	recordset.Open vQuerySQL,conn
	conn.Close

	AddVerificationInfoToResult "MSISDN",DictionaryTest_G.Item("ACCNTMSISDN")
	AddVerificationInfoToResult "AccountNo",DictionaryTest_G("AccountNo")
    AddVerificationInfoToResult "RootProduct",DictionaryTest_G(sRootProduct)
	
	
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : RetrieveAccountQTP_fn()
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveAccountQTP_fn()

'	strSQL = "SELECT * FROM [ExecuteDBQuery$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Database.xls",strSQL)
'	sDB = adoData( "DB")  & ""
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sDataDefinition = adoData("DataDefinition") & ""
    
	sAccnt = adoData("Account") & ""
	sMSISDN = adoData("MSISDN") & ""
	sRootProduct = adoData("RootProduct") & ""

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
	sQuery = "SELECT QUERY_SQL, QUERY_TYPE, PROMOTION_KEY FROM MINING_QUERY WHERE ENVIRONMENT='" & sEnv &"'  AND QUERY_NAME='" & sDataDefinition & "'"

    Set recordset = CreateObject("ADODB.Recordset")
	recordset.Open sQuery,conn
	i=0
	Do While Not recordset.EOF			
		vQuerySQL = recordset.Fields.Item(0)
		vQueryType = recordset.Fields.Item(1)
		vPromotionKey = recordset.Fields.Item(2)
		i = i+1
		Exit Do
	Loop

	 If i=0 Then
		AddVerificationInfoToResult  "Fail" , "QUERY_NAME not found"
		iModuleFail = 1
	End If
	On error resume next
	recordset.Close
	err.clear
	conn.Close

	If vPromotionKey<>"" Then

	Call RetrievePromotionDetailsQTP1_fn(vPromotionKey)



'		Dim dictProduct
'		Set dictProduct = CreateObject("Scripting.Dictionary")
'		vOutput=vDataProxyConnector.TryRetrieveProductData("15.2",vPromotionKey, dictProduct)
'		If (vOutput <> "true") Then
'	'	If (vDataProxyConnector.TryRetrieveProductData("15.2",sPromotion, dictProduct) = False) Then
'			 iModuleFail = 1
'			AddVerificationInfoToResult "Error",vOutput & " Promotion  key not retrieved"
'			Exit Function
'	
'		End If
		'DictionaryTest_G.Item("ProductName")= dictProduct("PRODUCT_NAME")
		'AddVerificationInfoToResult "Promotion Name",dictProduct("PRODUCT_NAME")
		
		If dictProduct.Exists ("RootProduct0") Then
			DictionaryTest_G("RootProduct0") = DictionaryTest_G.Item("RootProduct0")
			DictionaryTempTest_G("RootProduct0") = DictionaryTest_G.Item("RootProduct0")
			'DictionaryTest_G(sRootProduct) = dictProduct("ROOT_PRODUCT")
			'DictionaryTempTest_G(sRootProduct) = dictProduct("ROOT_PRODUCT")
			vQuerySQL = Replace(vQuerySQL,"RootProduct",DictionaryTest_G.Item("RootProduct0"))
		End If
		vQuerySQL = vQuerySQL + " and exists (select 1 from s_asset sa, s_prod_int spi where spi.name = '" + DictionaryTest_G.Item("ProductName") + "' and spi.row_id = sa.prod_id and sa.owner_accnt_id = soe.row_id and sip.row_id = sa.bill_profile_id)"
'		vQuerySQL = vQuerySQL + "AND sa2.serial_num NOT BETWEEN '447000000000'AND '447099999999'"

				vQuerySQL = Replace (vQuerySQL,"and soex.X_VF_CREDIT_CLASS is not null","")
                vQuerySQL = Replace (vQuerySQL,"and sip.CC_NUMBER is null","")
                vQuerySQL = Replace (vQuerySQL,"and sip.BANK_AUTHOR_FLG = 'N'","")
                vQuerySQL = Replace (vQuerySQL,"and sip.CRDT_CRD_EXPT_DT is null","")
                vQuerySQL = Replace (vQuerySQL,"and sip.CRDT_CRD_EXP_MO_CD is null","")
                vQuerySQL = Replace (vQuerySQL,"and sip.CRDT_CRD_EXP_YR_CD is null","")
                vQuerySQL = Replace (vQuerySQL,"and sip.X_CC_ADDR is null","")
                vQuerySQL = Replace (vQuerySQL,"and sip.X_CC_POSTCODE is null","")
                vQuerySQL = Replace (vQuerySQL,"and sip.CCV_NUMBER is null","")
				
				vQuerySQL = Replace (vQuerySQL,"and soex.X_VF_CREDIT_CLASS is null","")
                vQuerySQL = Replace (vQuerySQL,"and soex.X_VET_REF_NO is null","")
				vQuerySQL = Replace (vQuerySQL,"and soex.X_VF_OUTCOME_X is null","")
                vQuerySQL = Replace (vQuerySQL,"and soex.X_VF_VET_RES_EXP_DATE is null","")
				
				vQuerySQL = Replace (vQuerySQL,"and sc.email_addr like '%@sqcmail.uk'","")
				

		 
		 
		 
	End If
		vQuerySQL = vQuerySQL +" AND ( ( soe.url NOT LIKE '#B') OR ( soe.url IS NULL ) OR ( soe.url = 'Corrupt'))"
		vQuerySQL = Replace(vQuerySQL,"SELECT soe.row_id", "SELECT soe.x_vf_customer_code CUSTOMER_CODE,sa2.SERIAL_NUM MSISDN")

    vQuerySQL = vQuerySQL + " and sa2.created > sysdate-40 and rownum = 1 order by soe.created desc "
	AddVerificationInfoToResult "Query",vQuerySQL

	strConn=ConnectionDetails(sEnv)
	comm.CommandTimeout=120
'	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	conn.Open strConn
	strError = conn.State
		
	If strError <>1 Then

		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established" & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If
	recordset.Open vQuerySQL,conn
	i=0
	Do While Not recordset.EOF
		DictionaryTest_G("AccountNo") = recordset.Fields.Item(0)
		DictionaryTest_G("ACCNTMSISDN") = recordset.Fields.Item(1)
		i=i+1
		Exit Do
    Loop
	On error resume next
	recordset.Close
	err.clear
	If i=0 Then
		AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
		iModuleFail = 1
		conn.Close
		Exit Function
	End If

'	vQuerySQL = "SELECT soe.x_vf_customer_code, sa2.SERIAL_NUM from s_org_ext soe, s_asset sa2 where sa2.owner_accnt_id = soe.row_id and sa2.SERIAL_NUM like ('44%') and soe.row_id ='" &rowId& "'"
'	vQuery = "select distinct z.service_num from s_org_ext soe , s_order o, s_order_item z where o.customer_id = soe.row_id and o.row_id = z.order_id and soe.x_vf_customer_code in ('"+DictionaryTest_G("AccountNo")+"') and z.service_num like '4%'"
'	recordset.Open vQuery,conn
'	i=0
'	Do While Not recordset.EOF
''		DictionaryTest_G("AccountNo") = recordset.Fields.Item(0)
'		DictionaryTest_G("ACCNTMSISDN") = recordset.Fields.Item(0)
'		i=i+1
'		recordset.MoveNext
''		Exit Do
'    Loop

'	On error resume next
'	recordset.Close
'	err.clear
'	conn.Close

'    If i=0 Then
'		AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
'		iModuleFail = 1
'		conn.Close
'		Exit Function
'	End If
	vQuerySQL = "update s_org_ext soe set soe.url='#B' where soe.x_vf_customer_code ='"+DictionaryTest_G("AccountNo")+"'"
	recordset.Open vQuerySQL,conn
	conn.Close

	AddVerificationInfoToResult "MSISDN",DictionaryTest_G.Item("ACCNTMSISDN")
	AddVerificationInfoToResult "AccountNo",DictionaryTest_G("AccountNo")
    AddVerificationInfoToResult "RootProduct",DictionaryTest_G(sRootProduct)
	
	
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : RetrieveAccountQTPSecondTime_fnold()
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveAccountQTPSecondTime_fnold()

'	strSQL = "SELECT * FROM [ExecuteDBQuery$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Database.xls",strSQL)
'	sDB = adoData( "DB")  & ""
	Dim adoData	  
    strSQL = "SELECT * FROM [RetrieveAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sDataDefinition = adoData("DataDefinition") & ""
	sAccnt = adoData("Account") & ""
	sMSISDN = adoData("MSISDN") & ""
	sRootProduct = adoData("RootProduct") & ""

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
	sQuery = "SELECT QUERY_SQL, QUERY_TYPE, PROMOTION_KEY FROM MINING_QUERY WHERE ENVIRONMENT='" & sEnv &"'  AND QUERY_NAME='" & sDataDefinition & "'"
    Set recordset = CreateObject("ADODB.Recordset")
	recordset.Open sQuery,conn
	i=0
	Do While Not recordset.EOF			
		vQuerySQL = recordset.Fields.Item(0)
		vQueryType = recordset.Fields.Item(1)
		vPromotionKey = recordset.Fields.Item(2)
		i = i+1
		Exit Do
	Loop

	 If i=0 Then
		AddVerificationInfoToResult  "Fail" , "QUERY_NAME not found"
		iModuleFail = 1
	End If
	On error resume next
	recordset.Close
	err.clear
	conn.Close

	If vPromotionKey<>"" Then
		Dim dictProduct
		Set dictProduct = CreateObject("Scripting.Dictionary")
		vOutput=vDataProxyConnector.TryRetrieveProductData("15.2",vPromotionKey, dictProduct)
		If (vOutput <> "true") Then
	'	If (vDataProxyConnector.TryRetrieveProductData("15.2",sPromotion, dictProduct) = False) Then
			 iModuleFail = 1
			AddVerificationInfoToResult "Error",vOutput & " Promotion  key not retrieved"
			Exit Function
	
		End If
		DictionaryTest_G("New ProductName")= dictProduct("PRODUCT_NAME")
		AddVerificationInfoToResult "New Promotion Name",dictProduct("PRODUCT_NAME")
		If dictProduct.Exists ("ROOT_PRODUCT") Then
			DictionaryTest_G("RootProductNew") = dictProduct("ROOT_PRODUCT")
			DictionaryTempTest_G("RootProductNew") = dictProduct("ROOT_PRODUCT")
			vQuerySQL = Replace(vQuerySQL,"RootProduct",dictProduct ("ROOT_PRODUCT"))
		End If
		vQuerySQL = vQuerySQL + " and exists (select 1 from s_asset sa, s_prod_int spi where spi.name = '" + dictProduct("PRODUCT_NAME") + "' and spi.row_id = sa.prod_id and sa.owner_accnt_id = soe.row_id and sa.status_cd = 'Active' and sip.row_id = sa.bill_profile_id)"
'		vQuerySQL = vQuerySQL + "AND sa2.serial_num NOT BETWEEN '447000000000'AND '447099999999'"
	
		 
	End If
		vQuerySQL = vQuerySQL +" AND ( ( soe.url NOT LIKE '#B') OR ( soe.url IS NULL ) OR ( soe.url = 'Corrupt'))"
		vQuerySQL = Replace(vQuerySQL,"SELECT soe.row_id", "SELECT soe.x_vf_customer_code CUSTOMER_CODE,sa2.SERIAL_NUM MSISDN")

    vQuerySQL = vQuerySQL + " and sa2.created > sysdate-40 and rownum = 1 order by soe.created desc "
	AddVerificationInfoToResult "Query",vQuerySQL

	strConn=ConnectionDetails(sEnv)
	comm.CommandTimeout=120
'	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	conn.Open strConn
	strError = conn.State
		
	If strError <>1 Then

		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established" & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If
	recordset.Open vQuerySQL,conn
	i=0
	Do While Not recordset.EOF
		DictionaryTest_G("NewAccountNo") = recordset.Fields.Item(0)
'		DictionaryTest_G("NEWACCNTMSISDN") = recordset.Fields.Item(1)
		i=i+1
		Exit Do
    Loop
	On error resume next
	recordset.Close
	err.clear
	If i=0 Then
		AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
		iModuleFail = 1
		conn.Close
		Exit Function
	End If

'	vQuerySQL = "SELECT soe.x_vf_customer_code, sa2.SERIAL_NUM from s_org_ext soe, s_asset sa2 where sa2.owner_accnt_id = soe.row_id and sa2.SERIAL_NUM like ('44%') and soe.row_id ='" &rowId& "'"
'	vQuery = "select distinct z.service_num from s_org_ext soe , s_order o, s_order_item z where o.customer_id = soe.row_id and o.row_id = z.order_id and soe.x_vf_customer_code in ('"+DictionaryTest_G("AccountNo")+"') and z.service_num like '4%'"
'	recordset.Open vQuery,conn
'	i=0
'	Do While Not recordset.EOF
''		DictionaryTest_G("AccountNo") = recordset.Fields.Item(0)
'		DictionaryTest_G("NEWACCNTMSISDN") = recordset.Fields.Item(0)
'		i=i+1
'		recordset.MoveNext
'	'	Exit Do
'    Loop

'	On error resume next
'	recordset.Close
'	err.clear
'	conn.Close

 '   If i=0 Then
'		AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
'		iModuleFail = 1
'		conn.Close
'		Exit Function
'	End If
	vQuerySQL = "update s_org_ext soe set soe.url='#B' where soe.x_vf_customer_code ='"+DictionaryTest_G("AccountNo")+"'"
	recordset.Open vQuerySQL,conn
	conn.Close

	AddVerificationInfoToResult "New MSISDN",DictionaryTest_G.Item("NEWACCNTMSISDN")
	AddVerificationInfoToResult "New AccountNo",DictionaryTest_G("NewAccountNo")
    AddVerificationInfoToResult "New RootProduct",DictionaryTest_G("RootProductNew")
	
	
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : RetrieveAccountQTPSecondTime_fnold()
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveAccountQTPSecondTime_fn()

                Dim adoData        
				strSQL = "SELECT * FROM [RetrieveAccount$] WHERE  [RowNo]=" & iRowNo
                Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
                sDataDefinition = adoData("DataDefinition") & ""
                sAccnt = adoData("Account") & ""
                sMSISDN = adoData("MSISDN") & ""
                sRootProduct = adoData("RootProduct") & ""

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
                sQuery = "SELECT QUERY_SQL, QUERY_TYPE, PROMOTION_KEY FROM MINING_QUERY WHERE ENVIRONMENT='" & sEnv &"'  AND QUERY_NAME='" & sDataDefinition & "'"
				 Set recordset = CreateObject("ADODB.Recordset")
                recordset.Open sQuery,conn
                i=0
                Do While Not recordset.EOF                                       
                                vQuerySQL = recordset.Fields.Item(0)
                                vQueryType = recordset.Fields.Item(1)
                                vPromotionKey = recordset.Fields.Item(2)
                                i = i+1
                                Exit Do
                Loop

                If i=0 Then
                                AddVerificationInfoToResult  "Fail" , "QUERY_NAME not found"
                                iModuleFail = 1
                End If
                On error resume next
                recordset.Close
                err.clear
                conn.Close

                If vPromotionKey<>"" Then
                Call RetrievePromotionDetailsQTP1_fn(vPromotionKey)
'                               Dim dictProduct
'                               Set dictProduct = CreateObject("Scripting.Dictionary")
'                               vOutput=vDataProxyConnector.TryRetrieveProductData("15.2",vPromotionKey, dictProduct)
'                               If (vOutput <> "true") Then
'               '               If (vDataProxyConnector.TryRetrieveProductData("15.2",sPromotion, dictProduct) = False) Then
'                                               iModuleFail = 1
'                                               AddVerificationInfoToResult "Error",vOutput & " Promotion  key not retrieved"
'                                               Exit Function
'               
'                               End If
'                               DictionaryTest_G("New ProductName")= dictProduct("PRODUCT_NAME")
'                               AddVerificationInfoToResult "New Promotion Name",dictProduct("PRODUCT_NAME")
                                If dictProduct.Exists ("RootProduct0") Then
										DictionaryTest_G("RootProductNew") = DictionaryTest_G.Item("RootProduct0")
										DictionaryTempTest_G("RootProductNew") = DictionaryTest_G.Item("RootProduct0")
                                '               DictionaryTest_G("RootProductNew") = dictProduct("ROOT_PRODUCT")
                                '               DictionaryTempTest_G("RootProductNew") = dictProduct("ROOT_PRODUCT")
										vQuerySQL = Replace(vQuerySQL,"RootProduct",DictionaryTest_G.Item("RootProductNew"))
                                End If
                                vQuerySQL = vQuerySQL + " and exists (select 1 from s_asset sa, s_prod_int spi where spi.name = '" + DictionaryTest_G.Item("ProductName") + "' and spi.row_id = sa.prod_id and sa.owner_accnt_id = soe.row_id and sa.status_cd = 'Active' and sip.row_id = sa.bill_profile_id)"
'                               vQuerySQL = vQuerySQL + "AND sa2.serial_num NOT BETWEEN '447000000000'AND '447099999999'"
                
                                
                End If
                                vQuerySQL = vQuerySQL +" AND ( ( soe.url NOT LIKE '#B') OR ( soe.url IS NULL ) OR ( soe.url = 'Corrupt'))"
                                vQuerySQL = Replace(vQuerySQL,"SELECT soe.row_id", "SELECT soe.x_vf_customer_code CUSTOMER_CODE,sa2.SERIAL_NUM MSISDN")

				vQuerySQL = vQuerySQL + " and sa2.created > sysdate-40 and rownum = 1 order by soe.created desc "
                AddVerificationInfoToResult "Query",vQuerySQL

                strConn=ConnectionDetails(sEnv)
                comm.CommandTimeout=120
'               set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
                conn.Open strConn
                strError = conn.State
                                
                If strError <>1 Then

                                iModuleFail = 1
                                AddVerificationInfoToResult  "Error" , "Siebel Database connection not established" & err.description
                                Exit Function
                Else
                                AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
                End If
                recordset.Open vQuerySQL,conn
                i=0
                Do While Not recordset.EOF
                                DictionaryTest_G("NewAccountNo") = recordset.Fields.Item(0)
                               DictionaryTest_G("NEWACCNTMSISDN") = recordset.Fields.Item(1)
                                i=i+1
                                Exit Do
    Loop
                On error resume next
                recordset.Close
                err.clear
                If i=0 Then
                                AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
                                iModuleFail = 1
                                conn.Close
                                Exit Function
                End If

'               vQuerySQL = "SELECT soe.x_vf_customer_code, sa2.SERIAL_NUM from s_org_ext soe, s_asset sa2 where sa2.owner_accnt_id = soe.row_id and sa2.SERIAL_NUM like ('44%') and soe.row_id ='" &rowId& "'"
'               vQuery = "select distinct z.service_num from s_org_ext soe , s_order o, s_order_item z where o.customer_id = soe.row_id and o.row_id = z.order_id and soe.x_vf_customer_code in ('"+DictionaryTest_G("AccountNo")+"') and z.service_num like '4%'"
'               recordset.Open vQuery,conn
'               i=0
'               Do While Not recordset.EOF
''                              DictionaryTest_G("AccountNo") = recordset.Fields.Item(0)
'                               DictionaryTest_G("NEWACCNTMSISDN") = recordset.Fields.Item(0)
'                               i=i+1
'                               recordset.MoveNext
'               '               Exit Do
'    Loop

'               On error resume next
'               recordset.Close
'               err.clear
'               conn.Close

'   If i=0 Then
'                               AddVerificationInfoToResult  "Fail" , "There are not clean source entities for this scenario"
'                               iModuleFail = 1
'                               conn.Close
'                               Exit Function
'               End If
                vQuerySQL = "update s_org_ext soe set soe.url='#B' where soe.x_vf_customer_code ='"+DictionaryTest_G("NewAccountNo")+"'"
                recordset.Open vQuerySQL,conn
                conn.Close

                AddVerificationInfoToResult "New MSISDN",DictionaryTest_G.Item("NEWACCNTMSISDN")
                AddVerificationInfoToResult "New AccountNo",DictionaryTest_G("NewAccountNo")
    AddVerificationInfoToResult "New RootProduct",DictionaryTest_G("RootProductNew")
                
                
                If Err.Number <> 0 Then
                                iModuleFail = 1
                                AddVerificationInfoToResult  "Error" , Err.Description
                End If

End Function




'#################################################################################################
' 	Function Name : RetrievePromotionDetailsQTP_fn()
' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function RetrievePromotionDetailsQTP_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [RetrievePromotionDetails$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	sPromotion = adoData("Promotion") & ""
	sRootProduct = adoData("RootProduct") & ""

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

	sQuery = "SELECT PART_NUMBER FROM PRODUCT_DATA WHERE ENVIRONMENT='" & sEnv &"' AND KEY='" & sPromotion & "'"
    Set recordset = CreateObject("ADODB.Recordset")
	recordset.Open sQuery,conn
	i=0
	Do While Not recordset.EOF			
		vPartNo = recordset.Fields.Item(0)
		i = i+1
		Exit Do
	Loop

	 If i=0 Then
		AddVerificationInfoToResult  "Fail" , "Promotion KEY not found in QTP database."
		iModuleFail = 1
		Exit Function
	End If
	On error resume next
	recordset.Close
	err.clear
	conn.Close

		If DictionaryTest_G.Exists("PartNo") Then
			DictionaryTest_G.Item("PartNo")=vPartNo
		else
			DictionaryTest_G.add "PartNo",vPartNo
		End If

		AddVerificationInfoToResult "Part Number",DictionaryTest_G.Item("PartNo")

	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	comm.CommandTimeout=120
	conn.Open strConn
	strError = conn.State

	If strError <>1 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If

	If vPartNo<>"" Then
        sQuery = "SELECT NAME FROM S_PROD_INT WHERE PART_NUM='" & vPartNo & "'"
		    Set recordset = CreateObject("ADODB.Recordset")
			recordset.Open sQuery,conn
			i=0
			Do While Not recordset.EOF			
				vProductName = recordset.Fields.Item(0)
				i = i+1
				Exit Do
			Loop
		
			 If i=0 Then
				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
				iModuleFail = 1
				Exit Function
			End If
			On error resume next
			recordset.Close
			err.clear
			conn.Close
		 
	End If
	
		If DictionaryTest_G.Exists("ProductName") Then
			DictionaryTest_G.Item("ProductName")=vProductName
		else
			DictionaryTest_G.add "ProductName",vProductName
		End If

	   AddVerificationInfoToResult "ProductName",DictionaryTest_G.Item("ProductName")

	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	comm.CommandTimeout=120
	conn.Open strConn
	strError = conn.State

	If strError <>1 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If

       sQuery = "select spi2.name from s_prom_item spm, s_prod_int spi, s_prod_int spi2 where spm.promotion_id = spi.row_id and spm.prod_id = spi2.row_id and spi2.name in ('Mobile phone service','Mobile broadband service','EBU Sharer','PAYM','CBU Sharer','Fixed Service','One Net Fixed Data Service') and spm.dflt_qty = '1' and spi.part_num ='" & vPartNo & "'"
        	Set recordset = CreateObject("ADODB.Recordset")
			recordset.Open sQuery,conn
			i=0
			Do While Not recordset.EOF			
				vRootProduct = recordset.Fields.Item(0)
				i = i+1
				Exit Do
			Loop
		
			 If i=0 Then
				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
				iModuleFail = 1
			End If
			On error resume next
			recordset.Close
			err.clear
			conn.Close		        

		If DictionaryTest_G.Exists("RootProduct0") Then
			DictionaryTest_G.Item("RootProduct0")=vRootProduct
		else
			DictionaryTest_G.add "RootProduct0",vRootProduct
		End If

        AddVerificationInfoToResult "RootProduct",DictionaryTest_G.Item("RootProduct0")
	
	
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : RetrievePromotionDetailsQTP_fn()
' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function RetrievePromotionDetailsCatalogQTP_fn (vPartNo)



	If DictionaryTest_G.Exists("PartNo") Then
		DictionaryTest_G.Item("PartNo")=vPartNo
	else
		DictionaryTest_G.add "PartNo",vPartNo
	End If

	AddVerificationInfoToResult "Part Number",DictionaryTest_G.Item("PartNo")

	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	comm.CommandTimeout=120
	conn.Open strConn
	strError = conn.State

	If strError <>1 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If

	If vPartNo<>"" Then
        sQuery = "SELECT NAME FROM S_PROD_INT WHERE PART_NUM='" & vPartNo & "'"
		    Set recordset = CreateObject("ADODB.Recordset")
			recordset.Open sQuery,conn
			i=0
			Do While Not recordset.EOF			
				vProductName = recordset.Fields.Item(0)
				i = i+1
				Exit Do
			Loop
		
			 If i=0 Then
				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
				iModuleFail = 1
				Exit Function
			End If
			On error resume next
			recordset.Close
			err.clear
			conn.Close
		 
	End If
	
		If DictionaryTest_G.Exists("PRODUCT_NAME") Then
			DictionaryTest_G.Item("PRODUCT_NAME")=vProductName
		else
			DictionaryTest_G.add "PRODUCT_NAME",vProductName
		End If

	   AddVerificationInfoToResult "PRODUCT_NAME",DictionaryTest_G.Item("PRODUCT_NAME")

	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	comm.CommandTimeout=120
	conn.Open strConn
	strError = conn.State

	If strError <>1 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If

       sQuery = "select spi2.name from s_prom_item spm, s_prod_int spi, s_prod_int spi2 where spm.promotion_id = spi.row_id and spm.prod_id = spi2.row_id and spi2.name in ('Mobile phone service','Mobile broadband service','EBU Sharer','PAYM','CBU Sharer','Fixed Service','One Net Fixed Data Service') and spm.dflt_qty = '1' and spi.part_num ='" & vPartNo & "'"
        	Set recordset = CreateObject("ADODB.Recordset")
			recordset.Open sQuery,conn
			i=0
			Do While Not recordset.EOF			
				vRootProduct = recordset.Fields.Item(0)
				i = i+1
				Exit Do
			Loop
		
			 If i=0 Then
				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
				iModuleFail = 1
			End If
			On error resume next
			recordset.Close
			err.clear
			conn.Close		        

		If DictionaryTest_G.Exists("ROOT_PRODUCT") Then
			DictionaryTest_G.Item("ROOT_PRODUCT")=vRootProduct
		else
			DictionaryTest_G.add "ROOT_PRODUCT",vRootProduct
		End If

        AddVerificationInfoToResult "ROOT_PRODUCT",DictionaryTest_G.Item("ROOT_PRODUCT")
	
	
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



''#################################################################################################
'' 	Function Name : RetrievePromotionDetailsQTP_fn()
'' 	Description : 
''   Created By :  Tarun
''	Creation Date :        
''##################################################################################################
'Public Function RetrievePromotionDetailsCatalogQTP_fn()
'
'	Dim adoData	  
'    strSQL = "SELECT * FROM [RetrievePromotionDetails$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
'	sPromotion = adoData("Promotion") & ""
'	sRootProduct = adoData("RootProduct") & ""
'
'
'
'		If DictionaryTest_G.Exists("PartNo") Then
'			vPartNo=DictionaryTest_G.Item("PartNo")
'
'		End If
'
'		AddVerificationInfoToResult "Part Number",DictionaryTest_G.Item("PartNo")
'
'	strConn=ConnectionDetails(sEnv)
'	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
'	comm.CommandTimeout=120
'	conn.Open strConn
'	strError = conn.State
'
'	If strError <>1 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
'		Exit Function
'	Else
'		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
'	End If
'
'	If vPartNo<>"" Then
'        sQuery = "SELECT NAME FROM S_PROD_INT WHERE PART_NUM='" & vPartNo & "'"
'		    Set recordset = CreateObject("ADODB.Recordset")
'			recordset.Open sQuery,conn
'			i=0
'			Do While Not recordset.EOF			
'				vProductName = recordset.Fields.Item(0)
'				i = i+1
'				Exit Do
'			Loop
'		
'			 If i=0 Then
'				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
'				iModuleFail = 1
'				Exit Function
'			End If
'			On error resume next
'			recordset.Close
'			err.clear
'			conn.Close
'		 
'	End If
'	
'		If DictionaryTest_G.Exists("ProductName") Then
'			DictionaryTest_G.Item("ProductName")=vProductName
'		else
'			DictionaryTest_G.add "ProductName",vProductName
'		End If
'
'	   AddVerificationInfoToResult "ProductName",DictionaryTest_G.Item("ProductName")
'
'	strConn=ConnectionDetails(sEnv)
'	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
'	comm.CommandTimeout=120
'	conn.Open strConn
'	strError = conn.State
'
'	If strError <>1 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
'		Exit Function
'	Else
'		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
'	End If
'
'       sQuery = "select spi2.name from s_prom_item spm, s_prod_int spi, s_prod_int spi2 where spm.promotion_id = spi.row_id and spm.prod_id = spi2.row_id and spi2.name in ('Mobile phone service','Mobile broadband service','EBU Sharer','PAYM','CBU Sharer','Fixed Service','One Net Fixed Data Service') and spm.dflt_qty = '1' and spi.part_num ='" & vPartNo & "'"
'        	Set recordset = CreateObject("ADODB.Recordset")
'			recordset.Open sQuery,conn
'			i=0
'			Do While Not recordset.EOF			
'				vRootProduct = recordset.Fields.Item(0)
'				i = i+1
'				Exit Do
'			Loop
'		
'			 If i=0 Then
'				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
'				iModuleFail = 1
'			End If
'			On error resume next
'			recordset.Close
'			err.clear
'			conn.Close		        
'
'		If DictionaryTest_G.Exists("RootProduct0") Then
'			DictionaryTest_G.Item("RootProduct0")=vRootProduct
'		else
'			DictionaryTest_G.add "RootProduct0",vRootProduct
'		End If
'
'        AddVerificationInfoToResult "RootProduct",DictionaryTest_G.Item("RootProduct0")
'	
'	
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult  "Error" , Err.Description
'	End If
'
'End Function



'#################################################################################################
' 	Function Name : RetrievePromotionDetailsQTP1_fn()
' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function RetrievePromotionDetailsQTP1_fn(vPromotionKey)

'	Dim adoData	  
'    strSQL = "SELECT * FROM [RetrievePromotionDetails$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
'	sPromotion = adoData("Promotion") & ""
'	sRootProduct = adoData("RootProduct") & ""

	If vPromotionKey = "" Then
		AddVerificationInfoToResult "Error", "Promotion key is blank."
		iModuleFail  = 1
		Exit Function
	End If

	strConn=ConnectionDetails("QTP")
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	'comm.CommandTimeout=120
	conn.Open strConn
	strError = conn.State
		
	If strError <>1 Then

		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "QTP Database connection not established"  & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "QTP Database connection  established"
	End If

	sQuery = "SELECT PART_NUMBER FROM PRODUCT_DATA WHERE ENVIRONMENT='" & sEnv &"' AND KEY='" & vPromotionKey & "'"
    Set recordset = CreateObject("ADODB.Recordset")
	recordset.Open sQuery,conn
	i=0
	Do While Not recordset.EOF			
		vPartNo = recordset.Fields.Item(0)
		i = i+1
		Exit Do
	Loop

	 If i=0 Then
		AddVerificationInfoToResult  "Fail" , "Promotion KEY not found in QTP database."
		iModuleFail = 1
		Exit Function
	End If
	On error resume next
	recordset.Close
	err.clear
	conn.Close

		If DictionaryTest_G.Exists("PartNo") Then
			DictionaryTest_G.Item("PartNo")=vPartNo
		else
			DictionaryTest_G.add "PartNo",vPartNo
		End If

		AddVerificationInfoToResult "Part Number",DictionaryTest_G.Item("PartNo")

	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	comm.CommandTimeout=120
	conn.Open strConn
	strError = conn.State

	If strError <>1 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If

	If vPartNo<>"" Then
        sQuery = "SELECT NAME FROM S_PROD_INT WHERE PART_NUM='" & vPartNo & "'"
		    Set recordset = CreateObject("ADODB.Recordset")
			recordset.Open sQuery,conn
			i=0
			Do While Not recordset.EOF			
				vProductName = recordset.Fields.Item(0)
				i = i+1
				Exit Do
			Loop
		
			 If i=0 Then
				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
				iModuleFail = 1
				Exit Function
			End If
			On error resume next
			recordset.Close
			err.clear
			conn.Close
		 
	End If
	
		If DictionaryTest_G.Exists("ProductName") Then
			DictionaryTest_G.Item("ProductName")=vProductName
		else
			DictionaryTest_G.add "ProductName",vProductName
		End If

	   AddVerificationInfoToResult "ProductName",DictionaryTest_G.Item("ProductName")

	strConn=ConnectionDetails(sEnv)
	set conn = CreateObject("ADODB.Connection")  '1: CREATE CONNECTION
	comm.CommandTimeout=120
	conn.Open strConn
	strError = conn.State

	If strError <>1 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Siebel Database connection not established"  & err.description
		Exit Function
	Else
		AddVerificationInfoToResult  "Info" , "Siebel Database connection  established"
	End If

       sQuery = "select spi2.name from s_prom_item spm, s_prod_int spi, s_prod_int spi2 where spm.promotion_id = spi.row_id and spm.prod_id = spi2.row_id and spi2.name in ('Mobile phone service','Mobile broadband service','EBU Sharer','PAYM','CBU Sharer','Fixed Service','One Net Fixed Data Service') and spm.dflt_qty = '1' and spi.part_num ='" & vPartNo & "'"
        	Set recordset = CreateObject("ADODB.Recordset")
			recordset.Open sQuery,conn
			i=0
			Do While Not recordset.EOF			
				vROOT_PRODUCT = recordset.Fields.Item(0)
				i = i+1
				Exit Do
			Loop
		
			 If i=0 Then
				AddVerificationInfoToResult  "Fail" , "Part Number not found in Siebel DB."
				iModuleFail = 1
				Exit Function
			End If
			On error resume next
			recordset.Close
			err.clear
			conn.Close	

		If DictionaryTest_G.Exists("RootProduct0") Then
			DictionaryTest_G.Item("RootProduct0")=vROOT_PRODUCT
		else
			DictionaryTest_G.add "RootProduct0",vROOT_PRODUCT
		End If

        AddVerificationInfoToResult "RootProduct",DictionaryTest_G.Item("RootProduct0")
	
	
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function
