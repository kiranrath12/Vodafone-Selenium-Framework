
Public Function ConnectionDetails(Env)
   Select Case Env
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

	Do while Not adoData.Eof

	sQuery = adoData( "Query")  & ""
	'sEnv = adoData( "Env")  & ""
	sRecordsAffected = adoData( "RecordsAffected")  & ""
	If (instr(1,sQuery,"RowId",1)>=1) Then
		sQuery=Replace(sQuery,"RowId","'" & DictionaryTest_G.Item("RowId") & "'")
	End If
	if (instr(1,sQuery,"OrderNo",1)>=1) Then
		sQuery=Replace(sQuery,"OrderNo","'" & DictionaryTest_G.Item("OrderNo") & "'")
	End If
	if (instr(1,sQuery,"ACCNTMSISDN",1)>=1) Then
		sQuery=Replace(sQuery,"ACCNTMSISDN","'" & DictionaryTest_G.Item("ACCNTMSISDN") & "'")
	End If

	if (instr(1,sQuery,"AgreementID",1)>=1) Then
		sQuery=Replace(sQuery,"AgreementID","'" & DictionaryTest_G.Item("AgreementID") & "'")
	End If

		if (instr(1,sQuery,"MSISDN",1)>=1) Then
		sQuery=Replace(sQuery,"MSISDN","'" & DictionaryTest_G.Item("MSISDN") & "'")
	End If



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

'	conn.Execute(sQuery)
	Set recordset = conn.Execute (sQuery, RecordsAffected)
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

	adoData.movenext
	Loop
'	If Cint(sRecordsAffected) >=RecordsAffected and RecordsAffected <> 0 Then
		AddVerificationInfoToResult  "Pass" , RecordsAffected & " row(s) updated"
'	else
'		iModuleFail = 1
'		AddVerificationInfoToResult  "Error" , "Expected rows updated: " & sRecordsAffected &   " Actual rows updated: " & RecordsAffected
'	End If

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
