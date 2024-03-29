'#################################################################################################
' 	Function Name : CreateNewJob_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CreateNewJob_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewJob$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)
	sNewJob = adoData( "NewJob")  & ""
	sQuery = adoData( "Query")  & ""
	sGo = adoData( "Go")  & ""
	sSubmitJob = adoData( "SubmitJob")  & ""
	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

'	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebScreenViews("ScreenViews").Goto "Service Account List View","L3"
	If sNewJob ="Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Jobs").SiebButton("New").SiebClick False
	End If

	If sQuery ="Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Jobs").SiebButton("Query").SiebClick False
	End If

	Browser("index:=0").Page("index:=0").Sync
    CaptureSnapshot()
	Set obj = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Jobs")
'	obj.ActivateRow 0

	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		
		sValue = adoData( "Value")  & ""
		sValue = Replace (sValue, "JobId",DictionaryTempTest_G("JobId"))
		If sUIName <> "" Then
		
			UpdateSiebList obj,sUIName,sValue
            CaptureSnapshot()
			If instr (sUIName, "CaptureTextValue") > 0  Then
				DictionaryTempTest_G ("JobId")=sValue

			End If
		End If
		adoData.MoveNext
	Loop

	If sGo ="Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Jobs").SiebButton("Go").SiebClick False
	End If

	If sSubmitJob ="Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Jobs").SiebButton("Submit Job").SiebClick False
	End If

		CaptureSnapshot()
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : JobParameters_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function JobParameters_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [JobParameters$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)


	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

	
	

	Browser("index:=0").Page("index:=0").Sync
	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Job Parameters")
	objApplet.SiebList("micclass:=SiebList","repositoryname:=SiebList").ActivateRow 0
'	sRowId = DictionaryTest_G.Item("RowId")

	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
		If sValue = "[Id]='SRRowId'"  Then
			sValue = Replace (sValue, "SRRowId",DictionaryTempTest_G ("RowId"))
		ElseIf sValue = "[Id]='RowId'" Then
			sValue = Replace (sValue, "RowId",DictionaryTempTest_G ("RowId"))
		End If
		
		iIndex = 0
		
		If sLocateCol <> "" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If	
	
		If sUIName <> "" Then		
			UpdateSiebList objApplet,sUIName,sValue
            CaptureSnapshot()
		End If
	
		adoData.MoveNext
	Loop

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : NavigationToJobs_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function NavigationToJobs_fn()


	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")
	SiebApplication("Siebel Call Center").SiebToolbar("HIMain").Click "SiteMap"
	Browser("index:=0").Page("index:=0").Sync

'	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Administration - Server").Click

    Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("xpath:=//a[text()='Administration - Server Management']","index:=0").Click 
	Browser("index:=0").Page("index:=0").Sync
	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("xpath:=//a[text()='Jobs']","index:=0").Click
	Browser("index:=0").Page("index:=0").Sync
'	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Administration NewCo").Click
'	
'	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Jobs").Click


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : JobParameters_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SelectJob_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [SelectJob$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)
	sFind = adoData( "Find")  & ""
	sStartWith = adoData( "StartWith")  & ""

	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")
	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Components/Jobs").SiebPicklist("SiebPicklist").SiebSelect sFind
	
	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Components/Jobs").SiebText("SiebText").SiebSetText sStartWith
	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Components/Jobs").SiebButton("Go").SiebClick False
	If SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Components/Jobs").SiebButton("OK").Exist(2) Then
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Components/Jobs").SiebButton("OK").SiebClick False
	End If
	

	Browser("index:=0").Page("index:=0").Sync
	

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : NavigateToDeliveryManagement_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function NavigateToDeliveryManagement_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [CreateNewJob$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)
'	sNewJob = adoData( "NewJob")  & ""
	
	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebToolbar("HIMain").SiebClick "SiteMap"
	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Administration NewCo").Click
	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()
	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Delivery Method Validation").Click
	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : NavigationToEmployees_fn
' 	Description : This function is used to click on siete map and click on administratiion user and then click on emplyess
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function NavigationToEmployees_fn()


	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebToolbar("HIMain").Click "SiteMap"
	Browser("index:=0").Page("index:=0").Sync
	'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Administration - User").Click
	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("xpath:=//a[text()='Administration - User']","index:=0").Click
	
	Browser("index:=0").Page("index:=0").Sync
	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("xpath:=//a[text()='Employees']","index:=4").Click

	Browser("index:=0").Page("index:=0").Sync
'	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("xpath:=//a[text()='Employees List']","index:=0").Click

	CaptureSnapshot()



	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : AdministratorDivision_fn
' 	Description : This function is used to capture Division value 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AdministratorDivision_fn()

    strSQL = "SELECT * FROM [Division$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)


	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

	sUser = adoData( "User")  & ""
	sDivision = adoData( "Division")  & ""

	SiebApplication("Siebel Call Center").SiebScreen("Administration - User").SiebView("Employee Administration").SiebApplet("Employee").SiebButton("Query").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Administration - User").SiebView("Employee Administration").SiebApplet("Employee").SiebText("User ID").SiebSetText sUser
	SiebApplication("Siebel Call Center").SiebScreen("Administration - User").SiebView("Employee Administration").SiebApplet("Employee").SiebButton("Go").SiebClick False

'	sDivisionVal = SiebApplication("Siebel Call Center").SiebScreen("Administration - User").SiebView("Employee Administration").SiebApplet("Employee").SiebText("Division").GetRoProperty("text")
	sDivisionVal =  SiebApplication("Siebel Call Center").SiebScreen("Administration - User").SiebView("Employee Administration").SiebApplet("Employees").SiebList("List").SiebText("Position").GetRoProperty("text")

	If Trim(sDivisionVal) = sDivision Then
		AddVerificationInfoToResult  "Info" , "Division Value is " & sDivisionVal & " for " & sUser
	Else
		AddVerificationInfoToResult  "Info" , "Division Value is not populated correctly."
		iModuleFail = 1
	End If


	CaptureSnapshot()


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function
'#################################################################################################
' 	Function Name : GenerateContractTV_fn
' 	Description : 
'   Created By :  Pushkar 
'	Creation Date :        
'##################################################################################################
Public Function GenerateContractTV_fn()

	'Get Data
	Dim adoData	  
'    strSQL = "SELECT * FROM [CreateNewJob$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebMenu("Reports").Select "Reports\\Generate Contract - TV"

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Catalogue").SiebApplet("Output Type").SiebButton("Submit").Click

	Browser("index:=0").Page("index:=0").Sync

	If  Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(2) Then
		Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").WebClick
	End If

CaptureSnapshot()

	If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist Then
		Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
	End If



		If Browser("Certificate Error: Navigation").WinButton("To help protect your security,").Exist Then
			Browser("Certificate Error: Navigation").WinButton("To help protect your security,").Click
			Browser("Certificate Error: Navigation").WinMenu("menuobjtype:= 3 ").Select "Download File..."
		End If

CaptureSnapshot()

	Browser("Certificate Error: Navigation").Dialog("File Download").WinButton("Save").Click

	Browser("index:=0").Page("index:=0").Sync

	Dialog("Save As").WinToolbar("Address: C:\Users\ankit.lodha2").Type ("C:\TV reports") ''tcFolderPath

	Dialog("Save As").WinEdit("Edit").Set "test" '""'sTestCaseName

CaptureSnapshot()

	Dialog("Save As").WinButton("Save").Click

CaptureSnapshot()

'	Dialog("Download complete").WinButton("Close").Click

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Catalogue").SiebApplet("Output Type").SiebButton("Close").Click

	 'Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Close

'	Dialog("Save As").WinToolbar("ToolbarWindow32").Highlight
'	Dialog("Save As").WinToolbar("ToolbarWindow32").Press 1
'	Dialog("Save As").WinEdit("Edit_2").WinSet resFolder
'	Dialog("Save As").WinEdit("Edit_2").Type  micReturn 

	
End Function

'#################################################################################################
' 	Function Name : SetJobParameter_fn
' 	Description : This function fetches row id for required rows
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function SetJobParameter_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [SetJobParameter$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	Browser("index:=0").Page("index:=0").Sync

	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""

	Index = adoData( "Index")  & ""
	If Index="" Then
		Index=0
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Agreement List View","L3"
	Browser("index:=0").Page("index:=0").Sync
	'Flow
	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Administration - Server").SiebView("Component Requests").SiebApplet("Job Parameters")

			Do while Not adoData.Eof
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""
					iIndex = adoData( "Index")  & ""
					If iIndex="" Then
						iIndex=0
					End If
				If sLocateCol <> "" Then
					res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
					If  Cstr(res)<>"True" Then
						 iModuleFail = 1
						AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
						Exit Function
					End If
				End If	 
			If sUIName <> "" Then
					If instr (sUIName, "CaptureTextValue") > 0  Then
						sKey1=sValue	
					End If
					 UpdateSiebList objApplet,sUIName,sValue
					 If instr (sUIName, "CaptureTextValue") > 0  Then
						DictionaryTest_G(sKey1)=sValue
			End If
		End If
			adoData.MoveNext
			Loop ''

End Function



'#################################################################################################
' 	Function Name : NavigateToAdminData_fn
' 	Description : 
'   Created By :  Tarun Bansal
'	Creation Date :        
'##################################################################################################
Public Function NavigateToAdminData_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AdminData$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)
'	sNewJob = adoData( "NewJob")  & ""
	
	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebToolbar("HIMain").Click "SiteMap"
	Browser("index:=0").Page("index:=0").Sync

'	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Administration - Data").Click
		Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("xpath:=//a[text()='Administration - Data']","index:=0").Click
	Browser("index:=0").Page("index:=0").Sync
	
	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("List of Values").Click
	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values").SiebMenu("Menu").Select "ColumnsDisplayed"
	wait 3
	strArray = Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebList("HiddenItems").GetROProperty("all items")
	strArrayVal = Split(strArray,";")
	For each val in strArrayVal
	If val = "High" Then
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebList("HiddenItems").Select "High"
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").Image("ButtonShowItem").Click
		Exit For
	End If

	Next

	Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Save").Click

	wait 2
	
	SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values").SiebButton("Query").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values").SiebList("List").SiebText("Type").SiebSetText "VF_DPA_VALIDATION"
	SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values").SiebButton("Go").SiebClick False

		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values")

			Do while Not adoData.Eof
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""
				
					iIndex = adoData( "Index")  & ""
					If iIndex="" Then
						iIndex=0
					End If
				If sLocateCol <> "" Then
					res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
					If  Cstr(res)<>"True" Then
						 iModuleFail = 1
						AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
						Exit Function
					End If
				End If	 
				If sUIName <> "" Then
					 UpdateSiebList objApplet,sUIName,sValue
					CaptureSnapshot()
				End If
			adoData.MoveNext
			Loop ''


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : VerifyListOfValues_fn
' 	Description : This function fetches row id for required rows
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################
Public Function VerifyListOfValues_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [ListOfValues$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)

	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

	sNavigatetoListValue  = adoData( "NavigatetoListValue")  & ""
	sQuery  = adoData( "Query") & ""

			'Flow
		If sNavigatetoListValue = "Y" Then
		
		SiebApplication("Siebel Call Center").SiebToolbar("HIMain").Click "SiteMap"
		Browser("index:=0").Page("index:=0").Sync
		Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Administration - Data").Click
		Browser("index:=0").Page("index:=0").Sync
		Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("List of Values").Click
		Browser("index:=0").Page("index:=0").Sync

		End If

	If sQuery = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values").SiebButton("Query").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
	End if

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values")

	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
    	
		iIndex = 0
		
		If sLocateCol <> "" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			CaptureSnapshot()
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If	
	
		If sUIName <> "" Then		
			UpdateSiebList objApplet,sUIName,sValue
            CaptureSnapshot()
		End If
		
		adoData.MoveNext
	Loop

	   On error resume next
		SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values").SiebList("List").SiebText("Type").ProcessKey "EnterKey"
	err.clear
	CaptureSnapshot()

	If sNavigatetoListValue = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Administration - Data").SiebView("List of Values Administration").SiebApplet("List of Values").SiebButton("ToggleListRowCount").Click
		Browser("index:=0").Page("index:=0").Sync
	End If

		If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If	

 End Function



'#################################################################################################
' 	Function Name : NavigateToActivities_fn
' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function NavigateToActivities_fn()
	
	'sOR
	call SetObjRepository ("ServerManagement",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebToolbar("HIMain").Click "SiteMap"

	If Browser("Siebel Call Center").Page("Siebel Call Center_2").Frame("View Frame").Link("Activities").Exist(5) Then
		Browser("Siebel Call Center").Page("Siebel Call Center_2").Frame("View Frame").Link("Activities").Click
	Else
		AddVerificationInfoToResult  "Error" , "Site Map link is not clicked successfully."
		iModuleFail = 1
		Exit Function
	End If

	If Browser("Siebel Call Center").Page("Siebel Call Center_2").Frame("View Frame").Link("Activities_2").Exist(3) Then
		Browser("Siebel Call Center").Page("Siebel Call Center_2").Frame("View Frame").Link("Activities_2").Click
	Else
		AddVerificationInfoToResult  "Error" , "Activities link is not clicked successfully."
		iModuleFail =1
	End If

	Browser("index:=0").Page("index:=0").Sync
	If SiebApplication("Siebel Call Center").SiebScreen("Activities").SiebPDQ("PDQ").Exist(5) Then

		strArray = "1.All open items 14 days;2.All Activities;3.All open items;4.My items today"
		strSplit = Split(strArray,";")
		For loopVar = 0 to 3
			strVal = SiebApplication("Siebel Call Center").SiebScreen("Activities").SiebPDQ("PDQ").GetPDQByIndex(loopVar)
			If strSplit(loopVar) = strVal Then
				AddVerificationInfoToResult  "Info" , "Value : " & strVal & " is at position " & loopVar
			Else
				AddVerificationInfoToResult  "Info" , "Value : " & strVal & " is not at position " & loopVar
				iModuleFail = 1
			End If
		Next
	Else
		AddVerificationInfoToResult  "Error" ,"Saved Queries list page is not displayed."
		iModuleFail = 1
	End If
	CaptureSnapshot()


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function
