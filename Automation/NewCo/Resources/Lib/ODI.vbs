'#################################################################################################
' 	Function Name : LoginToODI_fn
' 	Description : This function log into ODI Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function LoginToODI_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [LoginToODI$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ODI.xls",strSQL)

	sURL = adoData("URL"&sEnv) & ""
	sUsername = adoData("Username") & ""
	sPassword = adoData("Password") & ""
	sRepository = adoData("Repository") & ""
	
	'sOR
	call SetObjRepository ("ODI",sProductDir & "Resources\")
    Call CloseAllBrowsers
	systemUtil.Run "IEXPLORE.EXE",sURL,"","Open",3

	Browser("Sign In").Page("Sign In").WebList("Repository").WebSelect sRepository
	Browser("Sign In").Page("Sign In").WebEdit("Username").WebSet sUsername
	Browser("Sign In").Page("Sign In").WebEdit("Password").WebSet sPassword
	Browser("Sign In").Page("Sign In").WebButton("Sign In").WebClick
'
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If
	CaptureSnapshot()

End Function

'#################################################################################################
' 	Function Name : BrowseScenarios_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function BrowseScenarios_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [BrowseScenarios$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ODI.xls",strSQL)

	sScenario = adoData("Scenario") & ""
	call SetObjRepository ("ODI",sProductDir & "Resources\")
	If Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("RuntimeExpand").Exist(0) Then
		Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("RuntimeExpand").WebClick
	End If

	If Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("ScenarioLoadPlanExpand").Exist(0) Then
		Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("ScenarioLoadPlanExpand").WebClick
	End If

	If Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("ScenariosExpand").Exist(0) Then
		Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("ScenariosExpand").WebClick
	End If

	Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("html tag:=SPAN","innertext:="&sScenario,"index:=0").WebClick
	Browser("Data Integrator Console").Page("Data Integrator Console").Image("Execute").WebClick

'
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

		CaptureSnapshot()

End Function


'#################################################################################################
' 	Function Name : ExecuteScenario_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ExecuteScenario_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [ExecuteScenario$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ODI.xls",strSQL)

	sAgent = adoData("Agent") & ""
	sContext = adoData("Context") & ""
	sLogLevel = adoData("LogLevel") & ""
	sDocStoreValue = adoData("DocStoreValue") & ""
	sSlot1 = adoData("Slot1") & ""
	sSlot2 = adoData("Slot2") & ""

	call SetObjRepository ("ODI",sProductDir & "Resources\")
	Browser("Data Integrator Console").Page("Data Integrator Console").WebList("Agent").WebSelect sAgent
	Browser("Data Integrator Console").Page("Data Integrator Console").WebList("Context").WebSelect sContext
	Browser("Data Integrator Console").Page("Data Integrator Console").WebList("LogLevel").WebSelect sLogLevel
	If sDocStoreValue = "Y" Then
		Browser("Data Integrator Console").Page("Data Integrator Console").WebEdit("shell:dialogRegion:1:exescen_v").WebSet sSlot1
		Browser("Data Integrator Console").Page("Data Integrator Console").WebEdit("shell:dialogRegion:1:exescen_v_2").WebSet sSlot2
		CaptureSnapshot()
	End If
	Browser("Data Integrator Console").Page("Data Integrator Console").WebButton("Execute Scenario").WebClick

	If Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("Scenario submitted").Exist Then
		scenario=Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("Scenario submitted").getROProperty("innertext")
		DictionaryTest_G("SessionId")=Replace(scenario,"Scenario submitted successfully. Created session with id :","")
		Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("DialogClose").WebClick
'		Browser("Data Integrator Console").Page("Data Integrator Console").WebButton("Cancel").WebClick
		CaptureSnapshot()
	else
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If
	CaptureSnapshot()

End Function


'#################################################################################################
' 	Function Name : SearchSessionStatus_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SearchSessionStatus_fn()

	call SetObjRepository ("ODI",sProductDir & "Resources\")

	sessionId = DictionaryTest_G("SessionId")
	Browser("Data Integrator Console").Page("Data Integrator Console").Link("Sessions").WebClick
	Browser("Data Integrator Console").Page("Data Integrator Console").WebEdit("SearchText").WebSet sessionId
	
	flag ="N"
	For i=0 to 10
		Browser("Data Integrator Console").Page("Data Integrator Console").WebButton("Search").WebClick

		If Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("Approved").Exist(2) Then
			AddVerificationInfoToResult  "Pass" , "Session is passed"
			flag ="Y"
			CaptureSnapshot()
			Exit For
		End If

		If Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("Expand Search Form").Exist(2) Then
			Browser("Data Integrator Console").Page("Data Integrator Console").WebElement("Expand Search Form").WebClick
			CaptureSnapshot()
		End If

		wait 5
	Next

	If flag ="N" Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Session is not passed"
		CaptureSnapshot()
	End If
	


	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If
	CaptureSnapshot()

End Function


'#################################################################################################
' 	Function Name : LoginToWCC_fn
' 	Description : This function log into ODI Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function LoginToWCC_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [LoginToWCC$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ODI.xls",strSQL)

	sURL = adoData("URL"&sEnv) & ""
	sUsername = adoData("Username") & ""
	sPassword = adoData("Password") & ""
	sRepository = adoData("Repository") & ""
	
	'sOR
	call SetObjRepository ("ODI",sProductDir & "Resources\")
    Call CloseAllBrowsers
	systemUtil.Run "IEXPLORE.EXE",sURL,"","Open",3

	Browser("Oracle WebCenter Content").Page("Oracle WebCenter Content").WebEdit("j_username").WebSet sUsername
	Browser("Oracle WebCenter Content").Page("Oracle WebCenter Content").WebEdit("j_password").WebSet sPassword
	Browser("Oracle WebCenter Content").Page("Oracle WebCenter Content").WebButton("Sign In").WebClick
'
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function

'#################################################################################################
' 	Function Name : Search_Invoice_WCC_fn
' 	Description : This function log into ODI Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Search_Invoice_WCC_fn()

	'DictionaryTest_G.Item("BillNo")="B1-364047"

	'sOR
	call SetObjRepository ("ODI",sProductDir & "Resources\")

	If Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").Link("Search").Exist(5) Then
		Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").Link("Search").WebClick
		Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").Link("Invoices").WebClick
		AddVerificationInfoToResult  "Info" ,"WCC Invoice Search is done successfully."
	Else
		AddVerificationInfoToResult  "Error" ,"WCC Login is not done successfully."
		iModuleFail = 1
		Exit Function
	End If

	If Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").WebEdit("xBillNo").Exist(5) Then
		Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").WebEdit("xBillNo").WebSet DictionaryTest_G.Item("BillNo")
		CaptureSnapshot()
		Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").WebButton("Search").WebClick
		AddVerificationInfoToResult  "Info" ,"Bill Number : " & DictionaryTest_G.Item("BillNo") & " is entered successfully."
	Else
		AddVerificationInfoToResult  "Error" ,"WCC Invoice Search is not done successfully."
		iModuleFail = 1
		Exit Function
	End If

	If Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").Link("ID").Exist(5) Then
		Browser("Oracle WebCenter Content").Page("Home Page for ohsadmin").Link("ID").WebClick
		AddVerificationInfoToResult  "Info" ,"WCC Invoice link is displayed successfully after entering Bill Number."
	Else
		AddVerificationInfoToResult  "Error" ,"WCC Invoice link is not displayed after entering Bill Number."
		iModuleFail = 1
		Exit Function
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function

