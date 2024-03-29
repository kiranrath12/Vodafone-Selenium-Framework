'#################################################################################################
' 	Function Name : ShowMoreButton_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ShowMoreButton_fn()

   call SetObjRepository ("Common",sProductDir & "Resources\")

	If Browser("title:=Siebel Call Center").Page("micclass:=Page").Frame("title:=Line Items","html tag:=FRAME").WebElement("innertext:=Menu Delete.*","html tag:=TR").Image("alt:=Show more","index:=0").Exist(10)  Then
		Browser("title:=Siebel Call Center").Page("micclass:=Page").Frame("title:=Line Items","html tag:=FRAME").WebElement("innertext:=Menu Delete.*","html tag:=TR").Image("alt:=Show more","index:=0").WebClick
'		ElseIf SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Catalogue").SiebApplet("Orders").SiebButton("ToggleLayout").Exist(1) Then
'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Catalogue").SiebApplet("Orders").SiebButton("ToggleLayout").SiebClick False
	End If
	


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : ShowMoreButton_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ShowMoreButtonUsedProductsServices_fn()

	If Browser("title:=Siebel Call Center").Page("micclass:=Page").Frame("title:=Account Summary").Image("xpath:=//td[contains(text(),'Used Product/Service')]/..//img[@title='Show more']").Exist(10)  Then
		Browser("title:=Siebel Call Center").Page("micclass:=Page").Frame("title:=Account Summary").Image("xpath:=//td[contains(text(),'Used Product/Service')]/..//img[@title='Show more']").WebClick
	End If
	


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : AboutRecordRowId_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function AboutRecordRowId_fn()
	call SetObjRepository ("Common",sProductDir & "Resources\")
	sRowId=Trim(Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("SiebWebPopupWindow").WebElement("RowId").WebElement("index:=1").GetROProperty("innertext"))
	AddVerificationInfoToResult  "Info" , "RowId is " & sRowId
	

	If DictionaryTest_G.Exists("RowId") Then
		DictionaryTest_G.Item("RowId")=sRowId
	Else
		DictionaryTest_G.Add "RowId",sRowId
	End If
    DictionaryTempTest_G ("RowId")=sRowId

Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("SiebWebPopupWindow").SblButton("OK").Click


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : CaptureDataOutputSheet_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CaptureDataOutputSheet_fn()
   	Set adoExcConn = CreateObject("ADODB.Connection")
    strSQL = "SELECT max([RowNo]) as RowNo FROM [CaptureDataOutputSheet$]"

	Set adoData = ExecExcQry(sProductDir & "Data\OutputSheet.xls",strSQL)

'	For each fld in adoData.Fields
'		msgbox fld.name
'
'	Next
'	DictionaryTest_G.Item("AccountNo")="abk"
'	DictionaryTest_G.Item("OrderNo")="giuksi"
'	DictionaryTest_G.Item("Status")="od"
		row=adoData("RowNo")
		strSQL="insert into  [CaptureDataOutputSheet$] ([RowNo],[Env],[TestName],[AccountNo],[OrderNo],[Status]) values("&row+1&",'"&sEnv&"','"&sTestCaseName&"','"&DictionaryTest_G.Item("AccountNo")&"','"&DictionaryTest_G.Item("OrderNo")&"','"&DictionaryTest_G.Item("Status")&"')"
		UpdateExcQry sProductDir & "Data\OutputSheet.xls",strSQL

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : RetrieveMSISDN_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveMSISDN_fn()
   	
'	strSQL = "SELECT TOP 1 * FROM [RetrieveMSISDN$] WHERE  [DataStatus] ='Y'"

	strSQL = "SELECT TOP 1 * FROM ["&sEnv&"$] WHERE  [DataStatus] ='Y'"
	'strSQL = "SELECT TOP 1 * FROM [RetrieveMSISDN$]"
   
	Set adoData = ExecExcQry(sProductDir & "Data\MSISDN.xls",strSQL)
	iRowNo = Cint(adoData( "RowNo"))

	sMSISDN = adoData( "MSISDN")  & ""
	sMSISDN = Replace(sMSISDN,"'","")

	sICCID = adoData( "ICCID")  & ""
	sICCID = Replace(sICCID,"'","")

	sDataStatus = adoData( "DataStatus")  & ""

'	   a = DictionaryTest_G.Keys   ' Get the keys.
'   For i = 0 To DictionaryTest_G.Count -1 ' Iterate the array.
'      print a(i)
'   Next

	If DictionaryTest_G.Exists("MSISDN") Then
		DictionaryTest_G.Item("MSISDN")=sMSISDN
	Else
		DictionaryTest_G.Add "MSISDN",sMSISDN
	End If
	

	AddVerificationInfoToResult "MSISDN",sMSISDN

	'DictionaryTest_G.add "MSISDN",sMSISDN

	If DictionaryTest_G.Exists("ICCID") Then
		DictionaryTest_G.Item("ICCID")=sICCID
	Else
		DictionaryTest_G.Add "ICCID",sICCID
	End If
	

	AddVerificationInfoToResult "ICCID",sICCID

	If DictionaryTest_G.Exists("DataStatus") Then
		DictionaryTest_G.Item("DataStatus")=sDataStatus
	Else
		DictionaryTest_G.Add "DataStatus",sDataStatus
	End If
'	DictionaryTest_G.add "sICCID",sICCID
'	DictionaryTest_G.add "DataStatus",sDataStatus

	strSQL = "update ["&sEnv&"$] set [DataStatus] ='N' where [RowNo]=" & iRowNo
	UpdateExcQry sProductDir & "Data\MSISDN.xls",strSQL

	'Set adoData = ExecExcQry(sProductDir & "Data\MSISDN.xls",strSQL)
'		strExcConn = GetExcConnStr(sProductDir & "Data\MSISDN.xls")
'		adoExcConn.Open strExcConn
'	adoExcConn.Execute strSQL
	'adoExcConn.Execute strSQL
'		Set adoData = ExecExcQry(sProductDir & "Data\MSISDN.xls",strSQL)
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : RetrieveMSISDNER_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveMSISDNER_fn()
   	
	strSQL = "SELECT TOP 1 * FROM [RetrieveMSISDNER$] WHERE  [DataStatus] ='Y'"
	'strSQL = "SELECT TOP 1 * FROM [RetrieveMSISDN$]"
   
	Set adoData = ExecExcQry(sProductDir & "Data\MSISDN.xls",strSQL)
	iRowNo = Cint(adoData( "RowNo"))
	sMSISDN = adoData( "MSISDN")  & ""
	sICCID = adoData( "ICCID")  & ""
	sDataStatus = adoData( "DataStatus")  & ""

'	   a = DictionaryTest_G.Keys   ' Get the keys.
'   For i = 0 To DictionaryTest_G.Count -1 ' Iterate the array.
'      print a(i)
'   Next

	If DictionaryTest_G.Exists("MSISDN") Then
		DictionaryTest_G.Item("MSISDN")=sMSISDN
	Else
		DictionaryTest_G.Add "MSISDN",sMSISDN
	End If
	'DictionaryTest_G.add "MSISDN",sMSISDN

	If DictionaryTest_G.Exists("sICCID") Then
		DictionaryTest_G.Item("sICCID")=sICCID
	Else
		DictionaryTest_G.Add "sICCID",sICCID
	End If

	If DictionaryTest_G.Exists("DataStatus") Then
		DictionaryTest_G.Item("DataStatus")=sDataStatus
	Else
		DictionaryTest_G.Add "DataStatus",sDataStatus
	End If
'	DictionaryTest_G.add "sICCID",sICCID
'	DictionaryTest_G.add "DataStatus",sDataStatus

	strSQL = "update [RetrieveMSISDN$] set [DataStatus] ='N' where [RowNo]=" & iRowNo
	UpdateExcQry sProductDir & "Data\MSISDN.xls",strSQL

	'Set adoData = ExecExcQry(sProductDir & "Data\MSISDN.xls",strSQL)
'		strExcConn = GetExcConnStr(sProductDir & "Data\MSISDN.xls")
'		adoExcConn.Open strExcConn
'	adoExcConn.Execute strSQL
	'adoExcConn.Execute strSQL
'		Set adoData = ExecExcQry(sProductDir & "Data\MSISDN.xls",strSQL)
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : CalculateETF_fn
' 	Description : 
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################
Public Function CalculateETF_fn()
   	
	strSQL = "SELECT TOP 1 * FROM [RetrieveMSISDNER$] WHERE  [DataStatus] ='Y'"
	Set adoData = ExecExcQry(sProductDir & "Data\MSISDN.xls",strSQL)


	sEndDate = DictionaryTest_G("EndDate") 
	sTodaysDate = Date

	End Function
'#################################################################################################
' 	Function Name : RetrieveCompRegNo_fn
' 	Description : 
'   Created By :  Zeba
'	Creation Date :        
'##################################################################################################
Public Function RetrieveCompRegNo_fn()
   	
'	strSQL = "SELECT TOP 1 * FROM [RetrieveMSISDN$] WHERE  [DataStatus] ='Y'"

	strSQL = "SELECT TOP 1 * FROM [CompRegNo$] WHERE  [DataStatus] ='Y'"
   
	Set adoData = ExecExcQry(sProductDir & "Data\CompRegNo.xls",strSQL)
	iRowNo = Cint(adoData( "RowNo"))

	sCompRegNo = adoData( "CompRegNo")  & ""
	sCompRegNo = Replace(sCompRegNo,"'","")

	sDataStatus = adoData( "DataStatus")  & ""

	If DictionaryTest_G.Exists("CompRegNo") Then
		DictionaryTest_G.Item("CompRegNo")=sCompRegNo
	Else
		DictionaryTest_G.Add "CompRegNo",sCompRegNo
	End If
	

	AddVerificationInfoToResult "CompRegNo",sCompRegNo

	If DictionaryTest_G.Exists("DataStatus") Then
		DictionaryTest_G.Item("DataStatus")=sDataStatus
	Else
		DictionaryTest_G.Add "DataStatus",sDataStatus
	End If

	strSQL = "update [CompRegNo$] set [DataStatus] ='N' where [RowNo]=" & iRowNo
	UpdateExcQry sProductDir & "Data\CompRegNo.xls",strSQL

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function