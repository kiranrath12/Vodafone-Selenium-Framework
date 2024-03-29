'#################################################################################################
' 	Function Name : LoginToSiebel_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function LoginToSiebel_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [LoginToSiebel$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)


	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sUsername = adoData("Username") & ""
	sPassword = adoData("Password") & ""


'		For i=1 to 2	
	Call CloseAllBrowsers


	
	'SystemUtil.Run "iexplore.exe",sURL
'	hwnd=Browser("CreationTime:=0").Object.HWND
'	If Window("hwnd:="&hwnd).getROProperty("maximized")<>true Then
'		Window("hwnd:="&hwnd).Maximize
'	End If

	Select Case sEnv
		Case "E4"
            sURL = "https://10.78.193.233:443/callcenter_enu/start.swe?SWECmd=AutoOn&SWEHo=10.78.193.233"
		Case "E7"
			sURL="https://10.78.195.105/callcenter_enu/start.swe?SWECmd=AutoOn"
		Case "C2"
			sURL = "https://10.78.221.37/callcenter_enu/start.swe?SWECmd=AutoOn&SWEHo=10.78.221.37"
			Case "SUP02"
			sURL = "https://10.78.199.133/callcenter_enu/start.swe?SWECmd=AutoOn"
			Case "E8"
			sURL = "https://10.78.195.233/callcenter_enu/start.swe?SWECmd=AutoOn"
			Case "E2"
			sURL = "https://10.78.193.41/callcenter_enu/start.swe?SWECmd=AutoOn"
			Case "SUP02"
			sURL = "https://10.78.199.133/callcenter_enu/start.swe?SWECmd=AutoOn"
	End Select

'	wait(10)

'	 sTitleMask = "QuickTest Professional - .*"
'   Window("regexpwndtitle:=" & sTitleMask).Maximize
'
'Sub MinimizeQTPWindow ()
'    Set     qtApp = getObject("","QuickTest.Application")
'    qtApp.WindowState = "Minimized"
'    Set qtApp = Nothing
'End Sub


'MinimizeQTPWindow
   If Window("Administrator: C:\Windows\Syst").Exist(2) Then
	Window("Administrator: C:\Windows\Syst").Minimize
	AddVerificationInfoToResult "Info" , "CMD minimized."
   End If
   wait(2)

	If Window("Server Manager_2").Exist(3) Then
		Window("Server Manager_2").Close
	 End If

	 wait 1


'   systemUtil.CloseProcessByName("mmc.exe")


	systemUtil.Run "IEXPLORE.EXE",sURL,"","Open",3

    'hwnd=Browser("CreationTime:=0").Object.HWND
    'hwnd=Browser("Certificate Error: Navigation").GetROProperty("HWND")
    'hwnd=Browser("CreationTime:=0").Object.HWND
'Window("hwnd:="&hwnd).Activate

  


    hwnd_Login = browser("name:=Certificate Error: Navigation Blocked").page("title:=Certificate Error: Navigation Blocked").GetROProperty("hwnd")
	Window("hwnd:="+cstr(hwnd_Login)).Activate
		AddVerificationInfoToResult "Info" , "browser activated."

	Window("hwnd:="+cstr(hwnd_Login)).click
		AddVerificationInfoToResult "Info" , "browser clicked"

	If  Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(2) Then
'		Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").highlight
		Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").WebClick
	End If

	If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist(3) Then
		Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
	End If


	If Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEUserName").Exist(3) Then
		Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEUserName").WebSet sUsername
		Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEPassword").WebSet sPassword
		Browser("Siebel Call Center").Page("Siebel Call Center").Image("Login").WebClick
'		AddVerificationInfoToResult "Info" , "login Page is displayed successfully."
		Else
		AddVerificationInfoToResult "Info" , "login Page is not displayed successfully."
'		iModuleFail = 1
	End If

		Browser("index:=0").Page("index:=0").Sync

	wait 7
	On error resume next
	Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("My Homepage").Click
	Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("My Homepage").Click
	Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("My Homepage").Click
	err.clear

	For i=1 to 3

		
		If SiebApplication("Siebel Call Center").Exist(7)=False Then
'			flag="N"
			iModuleFail = 0
			AddVerificationInfoToResult "Info" , "Trying for login again"
			Call CloseAllBrowsers
			systemUtil.Run "IEXPLORE.EXE",sURL,"","Open",3
		
		
			If  Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(2) Then
				Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").WebClick
			End If
		
			If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist(3) Then
				Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
			End If
		
		
			If Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEUserName").Exist(3) Then
				Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEUserName").WebSet sUsername
				Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEPassword").WebSet sPassword
				Browser("Siebel Call Center").Page("Siebel Call Center").Image("Login").WebClick
'				AddVerificationInfoToResult "Info" , "login Page is displayed successfully."
			Else
				AddVerificationInfoToResult "Info" , "login Page is not displayed successfully."
				iModuleFail = 1
			End If
			Wait 3
		else
			AddVerificationInfoToResult "Info" , "logged in successfully."
			Exit For
		End If
		
	Next

	If SiebApplication("Siebel Call Center").Exist(7)=False Then
		AddVerificationInfoToResult "Info" , "login Page is not displayed successfully."
		iModuleFail = 1
	End If

'   
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : LoginToSiebel_fn1
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function LoginToSiebel_fn1()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [LoginToSiebel$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sUsername = adoData("Username") & ""
	sPassword = adoData("Password") & ""
'		For i=1 to 2	
	Call CloseAllBrowsers

	systemUtil.Run "IEXPLORE.EXE",sURL,"","Open",3
	'SystemUtil.Run "iexplore.exe",sURL
'	hwnd=Browser("CreationTime:=0").Object.HWND
'	If Window("hwnd:="&hwnd).getROProperty("maximized")<>true Then
'		Window("hwnd:="&hwnd).Maximize
'	End If

	If  Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(2) Then
		Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").WebClick
	End If

	If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist Then
		Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
	End If

	If Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEUserName").Exist(3) Then
		Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEUserName").WebSet sUsername
		Browser("Siebel Call Center").Page("Siebel Call Center").WebEdit("SWEPassword").WebSet sPassword
		Browser("Siebel Call Center").Page("Siebel Call Center").Image("Login").WebClick
		AddVerificationInfoToResult "Info" , "login Page is displayed successfully."
		Else
		AddVerificationInfoToResult "Info" , "login Page is not displayed successfully."
		iModuleFail = 1
	End If

	wait 5

'    If Not(Browser("micclass:=Browser").Page("micclass:=Page").WaitProperty ("attribute/readyState", "complete", 10000)) Then
'	End If

'	If Browser("micclass:=Browser","name:=Siebel Call Center").Page("micclass:=Page","title:=Siebel Call Center").WebElement("micclass:=WebElement","innerhtml:=My Homepage").Exist Then
'	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").Exist Then
'		AddVerificationInfoToResult "Info" , "Account is login successfully."
'		Exit For
''	Else
''		AddVerificationInfoToResult "Fail" , "Account isnot  login successfully."
''		iModuleFail = 1
'	End If
'	Next
'
'	If Not(SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").Exist) Then
''		AddVerificationInfoToResult "Info" , "Account is login successfully."
''		Exit For
''	Else
'		AddVerificationInfoToResult "Fail" , "Account isnot  login successfully."
'		iModuleFail = 1
'	End If	
'   
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : TOOAplet_fn
' 	Description : This function performs Transfer of Ownership
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################

Public Function TOOAplet_fn()


	'Get Data
	Dim adoData	  
'    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
      
	sAccountNumber = DictionaryTest_G.Item("NewAccountNo")  '''This is done in accordance to RetriveAccountSecondTime_fn . Do not chnage
	
	'sOR

	call SetObjRepository ("Account",sProductDir & "Resources\")


	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Transfer of Ownership").SiebText("New Account").SiebSetText sAccountNumber
'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Transfer of Ownership").SiebButton("Continue").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Account").SiebList("List").SiebText("repositoryname:=Account Number","classname:=SiebText").SiebSetText sAccountNumber
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Transfer of Ownership").SiebText("repositoryname:=New Billing Profile","classname:=SiebText").OpenPopup
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing profile").SiebButton("OK").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Transfer of Ownership").SiebButton("Continue").SiebClick False

		CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : PromotionUpgrade_fn
' 	Description : This function performs Tariff Migration,Pre/Post
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function PromotionUpgrade_fn()


	'Get Data
	Dim adoData
	Dim sStartingWith

'    strSQL = "SELECT * FROM [PromotionUpgrade$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

'	sPromotionType = adoData("PromotionType") & ""
'	sStartingWith = adoData("StartingWith") & ""
	sStartingWith = DictionaryTest_G.Item("ProductName")
	sAction = DictionaryTest_G.Item("ActionforUsedProdService")

	Browser("index:=0").Page("index:=0").Sync

'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades").SiebPicklist("PromotionType").SiebSelect "Promotion Name"
'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades").SiebText("StartingWith").SiebSetText sStartingWith
'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades").SiebButton("Go").SiebClick False

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebPicklist("micclass:=SiebPicklist","repositoryname:=PopupQueryCombobox").SiebSelect "Promotion Name"
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebText("micclass:=SiebText","repositoryname:=PopupQuerySrchspec").SiebSetText sStartingWith
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebButton("micclass:=SiebButton","repositoryname:=PopupQueryExecute").SiebClick False
	CaptureSnapshot()

	Browser("index:=0").Page("index:=0").Sync


			If sAction = "Non like for like exchange" Then
					Set objApplet =SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Packages")
			 ElseIf sAction = "Tariff Migration" Then
					Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades")
			  ElseIf sAction = "Modify Promo" Then
					Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades ModifyPromo")
			End If			


    
	res=LocateColumns (objApplet,"Promotion Name",DictionaryTest_G.Item("ProductName"),"0")
	If  res=True Then
		AddVerificationInfoToResult  "Pass" , sLocateCol & "-" & sLocateColValue & " found in the list as expected"				 
'				Exit Function
	else
		iModuleFail = 1
		AddVerificationInfoToResult  "Fail" , sLocateCol & "-" & sLocateColValue & " not found in the list."
		
	End If

	rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebList("micclass:=SiebList","repositoryname:=SiebList").SiebListRowCount

	If rowCnt <> 0 Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebButton("micclass:=SiebButton","repositoryname:=Execute").SiebClick2 False
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Fail","No rows displayed after clicking on Go button."
		Exit Function
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
'               Function Name : CreateNewAccount_fn
'               Description : This function log into Siebel Application
'   Created By :  Ankit
'               Creation Date :        
'##################################################################################################
Public Function CreateNewAccount_fn()

                'Get Data
                Dim adoData        
    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	sTransact = adoData( "Transact")  & ""

	If sTransact="Y" Then
					sTitle=DictionaryTest_G.Item("Title")
					sFirst_Name =DictionaryTest_G.Item("Firstname")
					sLast_Name = DictionaryTest_G.Item("Surname")
					sBirth_Date = DictionaryTest_G.Item("DOB")
					sQASPostCode =DictionaryTest_G.Item("Postcode")
					sRegistrationNo = DictionaryTest_G.Item("COMP_REG_NUM")
	Else
					sFirst_Name = adoData( "First_Name")  & ""
					sLast_Name = adoData("Last_Name")  & ""
					sTitle=adoData("Title")  & ""
					sBirth_Date = adoData("Birth_Date")  & ""
					sQASPostCode = adoData("QASPostCode") & ""''Enter QAS postcode in the respective column
	End If

                
    sEmail = adoData("Email")  & ""
	sPost_Code = adoData("Post_Code")  & ""
	sAccount_Type=adoData("Account_Type")  & ""
	DictionaryTest_G.Add "sAccount_Type", sAccount_Type
	sAccount_SubCategory=adoData("AccountSubCategory")  & ""
	sAccount_Category=adoData("AccountCategory")  & ""
    sOnlineAccount=adoData("OnlineAccount")  & ""
    sPin1 = adoData("Pin1") & ""
	sPin2 = adoData("Pin2")
	sPin3 = adoData("Pin3")
	sPin4 = adoData("Pin4")
   sAddressLine=adoData("AddressLine")  & ""
   sUserName = adoData("User_name")  & ""
   sPostcode=adoData("Postcode")  & ""
'   sHouseNumber=adoData("HouseNumber")  & ""
    sSave=adoData("Save")  & ""  
	sCareTeam=adoData("CareTeam")  & ""
	sTradingAs = adoData("TradingAs")  & ""
	sAnonymous = adoData("Anonymous")  & ""
	sPhoneNumber = adoData("PhoneNumber")
	sEmail1 = adoData("Email1")
	sEmail2 = adoData("Email2")
	sAccountTab = adoData("AccountTab") & ""
	sNew = adoData("New") & ""
	sSkipAddress = adoData("AddressLine") & "" ''Should be "N" in account sheet to skip entering address
	sLegal = adoData("Legal Status") & ""
	sAnonymousDrillDown = adoData("AnonymousDrillDown") & ""
	sGender = adoData("Gender") & ""
	sPriorityFaultRepair = adoData("PriorityFaultRepair") & ""
    sMatchCompany = adoData("MatchCompany") & ""
	sRegistrationNo = adoData("RegistrationNo") & ""
	sAnonymousValidation = adoData("AnonymousValidation") & ""
	sAnonymousAddressValidation = adoData("AnonymousAddressValidation") & ""
	 sOnlineFlagValidation = adoData("OnlineFlagValidation")  & ""

     sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If	

	If sUserName<>"" Then

    Dim sDate : sDate = Day(Now)
	Dim sMonth : sMonth = Month(Now)
	Dim sYear : sYear = Year(Now)
	Dim sHour : sHour = Hour(Now)
	Dim sMinute : sMinute = Minute(Now)
	Dim sSecond : sSecond = Second(Now)
	
	'Create Random Number
	fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)
	sUserName=sUserName & fnRandomNumberWithDateTimeStamp
	End If
	
	Browser("index:=0").Page("index:=0").Sync

	If sAccountTab<>"No" Then
	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If
		
	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "PageTabs not present"
		Exit Function
	End If
	End If

	'To verify if account no, account type and acount status are blank before clicking new button

	If sNew <> "No" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("New").SiebClick False
	End If
	 ' clicking on New Button

		If sAnonymous = "Y"  Then ' Creating Anonymous Account
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("micclass:=SiebCheckbox","repositoryname:=Anonymous").SiebSetCheckbox sOnlineAccount
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account type").SiebSelect sAccount_Type
			If Ucase(sAccount_Category) <> "" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account Category").SiebSelect sAccount_Category
			End If
			If Ucase(sAccount_SubCategory) <> "" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account SubCategory").SiebSelect sAccount_SubCategory
			End If

			If sAnonymousAddressValidation = "Y" Then
			   SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup
			   Browser("index:=0").Page("index:=0").Sync
			   If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Addresses").Exist(4) Then
					If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebTable("OK").WebElement("innertext:=OK","index:=0","class:=miniBtnUICOff").Exist(5) Then
						AddVerificationInfoToResult  "Info" , "OK button is disabled for address applet for anonymous account which means we are not able to change the address."
						CaptureSnapshot()
						Exit Function
					Else
						AddVerificationInfoToResult  "Error" , "OK button is enabled for address applet for anonymous account but it should be disabled."
						iModuleFail = 1
						Exit Function
					End If
				Else
					AddVerificationInfoToResult  "Error" , "Address Applet is not clicked successfully."
					iModuleFail = 1
					Exit Function
					
			   End If
			End If
						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Save").SiebClick False
			AddVerificationInfoToResult  "Info" , "Anonymous Account is created successfully."
			


			If sAnonymousValidation = "Y" Then
				sCheckBox = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Create online account").GetRoProperty("isenabled")
				If sCheckBox = "False" Then
					AddVerificationInfoToResult  "Info" , "Create Online Account checkbox is disabled for Anonymous account as expected"
				Else
					AddVerificationInfoToResult  "Error" , "Create Online Account checkbox is enabled for Anonymous account."
					iModuleFail = 1
				End If

				sPostCode = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Postcode").GetRoProperty("text"))
					If sPostCode = "XX9 9XX" Then
						AddVerificationInfoToResult  "Info" , "Postcode is " & sPostCode & " for anonymous account as expected."
					Else
						AddVerificationInfoToResult  "Error" , "Postcode is coming out as " & sPostCode & " for anonymous account."
						iModuleFail = 1
					End If				
			End If

			CaptureSnapshot()
			If sAnonymousDrillDown= "N" Then
				Exit Function
			End If
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
			Exit Function
		End If

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account type").SiebSelect sAccount_Type


		If Ucase(sAccount_Category) <> "" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account Category").SiebSelect sAccount_Category
		End If
		If Ucase(sAccount_SubCategory) <> "" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account SubCategory").SiebSelect sAccount_SubCategory
		End If
	
    SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Title").SiebSelect sTitle
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("First name").SiebSetText sFirst_Name
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Last name").SiebSetText sLast_Name
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCalendar("Date of birth").SiebSetText sBirth_Date

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN1").SiebSetText sPin1
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN2").SiebSetText sPin2
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN3").SiebSetText sPin3
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN4").SiebSetText sPin4


	If sOnlineAccount = "On" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Update Email").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebText("Email").SiebSetText sEmail1
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebPicklist("SiebPicklist").SetText sEmail2
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebText("Re-type Email").SiebSetText sEmail1
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebPicklist("SiebPicklist_2").SetText sEmail2
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebButton("OK").SiebClick False


		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Create online account").SiebSetCheckbox sOnlineAccount
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Email").SiebSetText sEmail
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("micclass:=SiebText","repositoryname:=VF User Name").SiebSetText sUserName
	ElseIf sOnlineAccount = "Off" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Create online account").SiebSetCheckbox sOnlineAccount
	End If
     

		If Ucase(sPhoneNumber) <> "" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Alt phone no").SiebSetText sPhoneNumber
		End If
	If sSkipAddress <> "N" Then

			If  sPostcode = "QAS" Then

				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Postcode").SiebSetText sQASPostCode
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup
				Browser("index:=0").Page("index:=0").Sync
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("QAS").SiebButton("OK").Click
				Browser("index:=0").Page("index:=0").Sync

		   Else
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup
			   Browser("index:=0").Page("index:=0").Sync


			   	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SiebSetText sPostcode
				On error resume next
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").ProcessKey "EnterKey"
				err.clear


                Browser("index:=0").Page("index:=0").Sync

					Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address")
					res=LocateColumns (objApplet,"Address status","Validated","0")
						If  Cstr(res)<>"True" Then
							 iModuleFail = 1
							AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
							Exit Function
						End If

'			   	 If   Ucase(sPopUp)<>"FALSE" Then
'						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SiebSetText sPostcode
'						On error resume next
'						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").ProcessKey "EnterKey"
'						err.clear
'			 Else
'						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SiebSetText sPostcode
'						On error resume next
'						On error resume next
'						err.clear
'
'					Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address")
'					res=LocateColumns (objApplet,"Status","Validated","0")
'						If  Cstr(res)<>"True" Then
'							 iModuleFail = 1
'							AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
'							Exit Function
'						End If
'
'				End If
'
	End If

				
		
				If Ucase(sTradingAs) <> "" Then
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Trading As").SiebSetText sTradingAs
				End If


				If sPriorityFaultRepair = "Yes" Then
					'SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Priority Fault Repair").SetOn
					Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WinCheckBox("Button").Set "ON"
				End If


					If  sPostcode <> "QAS" Then
                        If sPopUp <> "NO" Then
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Addresses").SiebButton("Add >").SiebClick sPopUp
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Addresses").SiebButton("OK").SiebClick False
							Else
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Addresses").SiebButton("Add >").SiebClick sPopUp
							  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Addresses").SiebButton("OK").SiebClick False
					   End If
				   End If
				   End If
		

		If Ucase(sCareTeam) <> "" Then
						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Care Team").SiebSetText sCareTeam
				End If

	If Ucase(sAccount_SubCategory) <> "" Then
			sAccountName=sFirst_Name & " " & sLast_Name
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Account name").SiebSetText sAccountName
	End If

	sAddressLine = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Address Line 1").GetRoProperty("text"))

		If DictionaryTest_G.Exists("AddressLine") Then
			DictionaryTest_G.Item("AddressLine")=sAddressLine
		else
			DictionaryTest_G.add "AddressLine",sAddressLine
		End If

		If sLegal <> "" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Legal status").SiebSetText sLegal
		End If

		If sGender <> "" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Gender").SiebSelect sGender
		End If

		If sMatchCompany = "Y" Then
		
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Company reg. no.").Exist(2) Then
			  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Company reg. no.").SiebSetText sRegistrationNo
			Else 
				iModuleFail = 1
				AddVerificationInfoToResult "Fail","Company registration no. edit box is not present on the account page."
				Exit Function
			End If
		
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Match Company").SiebClick False
		
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Match Company").Exist(4) Then
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Match Company").SiebButton("OK").SiebClick False
			Else
				iModuleFail = 1
				AddVerificationInfoToResult "Fail","Match Company applet is not coming."
				Exit Function
			End If
		
		End If


	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Save").SiebClick False

	If sOnlineFlagValidation = "Y" Then
		sOnlineCheck = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Create online account").GetRoProperty("isenabled")
		If sOnlineCheck = "False" Then
			AddVerificationInfoToResult  "Info" , "As expected Online check box is disabled i.e not editable after saving account."
		Else
			AddVerificationInfoToResult  "Error" , "Online check box is enabled i.e editable after saving account."
			iModuleFail = 1
		End If
	End If

    CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : VerifyAccountSummary_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function VerifyAccountSummary_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)


	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow


	sFirst_Name = adoData( "First_Name")  & ""
	sLast_Name = adoData( "Last_Name")  & ""
'	sFirst_Name = DictionaryTest_G.Item("FirstName")
'	sLast_Name = DictionaryTest_G.Item("LastName")   
	 
	sAccountName= sFirst_Name & " " & sLast_Name
	Browser("index:=0").Page("index:=0").Sync
	actAccountName=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebGetCellText ("Account name", 0)
	actAccountNo=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebGetCellText ("Account no.", 0)

	If DictionaryTest_G.Exists("AccountNo") Then
		DictionaryTest_G.Item("AccountNo")=actAccountNo
	else
		DictionaryTest_G.add "AccountNo",actAccountNo
	End If

	AddVerificationInfoToResult  "Info" , "AccountNo is " & actAccountNo

	If Ucase(sAccountName)=Ucase(actAccountName) Then
		AddVerificationInfoToResult "Pass","Account name " & actAccountName & " is successfully created"
	else
		  iModuleFail = 1
		AddVerificationInfoToResult "Fail","Account creation is failed"
		Exit Function
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
	
	Browser("index:=0").Page("index:=0").Sync
'	actAccountName1= trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Account name").GetRoProperty("text"))
'	
'
'	If Ucase(sAccountName)=Ucase(actAccountName1) Then
'		AddVerificationInfoToResult "Pass","Account name " & actAccountName & " is displayed in Account summary page"
'	else
'		 iModuleFail = 1
'		AddVerificationInfoToResult "Fail","Account name " & actAccountName1 & " is not displayed in Account summary page"
'		Exit Function
'	End If

'   Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

'	Set objWsh = CreateObject("WScript.Shell")
'	objWsh.SendKeys ("{F9}")


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : ResetPin_fn
' 	Description : This functionis used to reset pin and memeorable word 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ResetPin_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [Reset$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sReset = adoData( "Reset")  & ""

	sPopUp = adoData( "Popup")  & ""


	sCreateNewCustComms = adoData( "CreateNewCustComms")  & ""

	If Ucase(sPopUp)="FALSE" or sPopUp=""Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If
	'Flow
	If  (sCreateNewCustComms = "Y" )Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=VF Account Activity List Applet - Summary").SiebButton("micclass:=SiebButton","repositoryname:=NewRecord").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=VF Account Activity List Applet - Summary").SiebList("micclass:=SiebList","repositoryname:=SiebList").SiebDrillDownColumn "ID",0
	End If
	

	If sReset = "Pin" Then
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Contacts").SiebButton("Set / Reset PIN").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset PIN").SiebText("PIN").SiebSetText "1"
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset PIN").SiebText("PIN_2").SiebSetText "2"
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset PIN").SiebText("PIN_3").SiebSetText "9"
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset PIN").SiebText("PIN_4").SiebSetText "9"
        CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset PIN").SiebButton("Ok").SiebClick sPopUp

	ElseIf sReset = "Memorable" Then
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Contacts").SiebButton("Set/ Reset Word and Hint").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset Word").SiebText("Memorable word").SiebSetText "automation12"
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset Word").SiebText("Memorable hint").SetText "automation14"
        CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Reset Word").SiebButton("Ok").SiebClick sPopUp

	ElseIf sReset = "MemorableWord" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Set/Reset Word and Hint").SiebClick False
		If  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Reset Word").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Reset Word").SiebText("Memorable word").SiebSetText "automation12"
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Reset Word").SiebText("Memorable hint").SiebSetText "automation"
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Reset Word").SiebButton("Ok").SiebClick sPopUp
				AddVerificationInfoToResult "Info","Set/Reset Word and Hint is done succesfully."
		Else
			AddVerificationInfoToResult "Fail","Set/Reset Word and Hint button is not clicked succesfully."
			iModuleFail = 1
		End If
	ElseIf sReset = "Password" Then
		'SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Contacts").SiebList("List").SiebText("Last name").SetText
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Contacts").SiebButton("Set / Reset Password").SiebClick sPopUp
		CaptureSnapshot()
		AddVerificationInfoToResult "Pass","Password is reset succesfully."

	ElseIf sReset = "AccountPermission" Then
		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Contacts")
			Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
'		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
'        sLocateColValue = Replace(sLocateColValue,"MSISDN",DictionaryTest_G.Item("ACCNTMSISDN"))
'		sUIName = adoData( "UIName")  & ""
'		sValue = adoData( "Value")  & ""
	iIndex = 0

'		sCollapse = adoData( "Collapse")  & ""
		
		If sLocateCol <> "" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			CaptureSnapshot()
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If	

'		If sUIName <> "" Then		
'			UpdateSiebList objApplet,sUIName,sValue
'			CaptureSnapshot()
'		End If

		CaptureSnapshot()
		adoData.MoveNext


			Loop
		End If
	 CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : CreateNewBillingProfile_fn
' 	Description : This function creates prepay billing profile in siebel
'   Created By :  Automation
'	Creation Date :        
'##################################################################################################
Public Function CreateNewBillingProfile_fn1()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewBillingProfile$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)


	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
	sPrePostPay = adoData( "PrePostPay")  & ""
	sPrimary = adoData( "Primary")  & ""


	'Flow	
	'*****************************************************************************
	'Check billing profile screen and create prepay profile

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").Exist Then
	'Create a billing Profile
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebButton("New").SiebClick False
	ElseIf SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").Exist Then
		 SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Billing Profiles").SiebButton("New").SiebClick False
	Else
	'Click on the account number link
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
	'Create a billing Profile
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebButton("New").SiebClick False
	End If

	strRowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebList("List").RowsCount

	If  strRowCount <> 0 Then
		AddVerificationInfoToResult  "Pass" , "Billing Profile created Successfully"
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebList("List").SiebPicklist("Postpay/Prepay").SiebSelect sPrePostPay
	Else
		iModuleFail = 1
		AddVerificationInfoToResult  "Fail" , "New Button is not clicked successfully."
	End If

	If sPrimary = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebList("List").SiebCheckbox("Primary").SiebSetCheckbox "ON"
	End If

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : CreateNewBillingProfile_fn
' 	Description : This function creates prepay billing profile in siebel
'   Created By :  Automation
'	Creation Date :        
'##################################################################################################
Public Function CreateNewBillingProfile_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewBillingProfile$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)


	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
'	sPrePostPay = adoData( "PrePostPay")  & ""
'	sPrimary = adoData( "Primary")  & ""
'	sTokenization = adoData( "Tokenization")  & ""
	sCreateViaProfiles = adoData( "CreateViaProfiles")  & ""
	sPaymentTermsDays = adoData( "PaymentTermsDays")  & ""
	sNewBillingProfile = adoData( "NewBillingProfile")  & ""

		sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

	'Flow

	'*****************************************************************************
	'Check billing profile screen and create prepay profile

	If SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Billing Profiles").Exist Then
	'Create a billing Profile
				If sCreateViaProfiles = "Y" Then
	
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3" ' clicking on Profiles tab
							Browser("index:=0").Page("index:=0").Sync
							CaptureSnapshot()
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile-Profiles").SiebButton("New").SiebClick False
							Browser("index:=0").Page("index:=0").Sync
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment terms (days)").SiebSelect sPaymentTermsDays
							DaysChanged = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment terms (days)").GetROProperty("activeitem")
							CaptureSnapshot()
								If Trim(DaysChanged) =  Trim(sPaymentTermsDays) Then
										AddVerificationInfoToResult "Info" , "Payment Terms Days changed successfully to" & sPaymentTermsDays
								Else
										AddVerificationInfoToResult "Info" , "Payment Terms Days not changed  to" & sPaymentTermsDays
										iModuleFail = 1
								End If
								Exit Function
					ElseIf sNewBillingProfile = "Y" Then
						SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Billing Profiles").SiebButton("New").SiebClick False
					ElseIf sNewBillingProfile = "INC" Then
						SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Billing Profiles").SiebButton("New").SiebClick sPopUp
					End If
		
	Else
	'Click on the account number link
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
	'Create a billing Profile
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Billing Profiles").SiebButton("New").SiebClick False
	End If


	''''''++++++++++++++++++++++++++++++NewCode+++++++++++++++++++++++++++++++++++++++

		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles")
			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

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

			If sNewBillingProfile = "RealTimeBalance" Then
				If SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Product / Services").SiebButton("Top-Up Request").Exist(5) Then
					AddVerificationInfoToResult  "Info" ,"Real Time Balance page is displayed directly after drill down on Name column in Billing Profile applet."
					CaptureSnapshot()
					Exit Function
				Else
					AddVerificationInfoToResult  "Error" ,"Real Time Balance page is not displayed directly after drill down on Name column in Billing Profile applet."
					iModuleFail = 1
					Exit Function
				End If
			End If

		strRowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebList("List").RowsCount

		If  strRowCount <> 0 Then
			AddVerificationInfoToResult  "Pass" , "Billing Profile created Successfully"
		Else
			iModuleFail = 1
			AddVerificationInfoToResult  "Fail" , "New Button is not clicked successfully."
		End If

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : MenuBillingProfile_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function MenuBillingProfile_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [MenuBillingProfile$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow


	sMenuItem = adoData( "MenuItem")  & ""



	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebMenu("Menu").SiebSelectMenu sMenuItem


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : CreateNewOrder_fn
' 	Description : This function creates new order and opens it
'   Created By :  Automation
'	Creation Date :        
'##################################################################################################

Public Function CreateNewOrder_fn()

    	'sOR
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewOrder$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	call SetObjRepository ("Account",sProductDir & "Resources\")

	sNewButton = adoData( "NewButton")  & ""

'	sPopup = "An address change has been processed on this account within the past 7 days"

		sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If 
	
'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "SIS OM Customer Account Portal View","L3" ' clicking on Accounts Screen tab

	Browser("index:=0").Page("index:=0").Sync

	If Ucase(sNewButton) <> "NO" Then

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebButton("New").SiebClick sPopUp
	
		If Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").Exist(3) Then
			Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").WebButton("Confirm").Click
		End If

			REM If 	Window("Address Change Alert --").Exist(2) Then
				REM Window("Address Change Alert --").Activate
				REM Window("Address Change Alert --").WinObject("Internet Explorer_Server").Click 342,166
			REM End If
	
	'	Wait(7)
		Browser("index:=0").Page("index:=0").Sync
	else
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").ActivateRow 0
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebDrillDownColumn "Order no.",0
	
	End If

	'DataTable.Value("Order_Number") = strOrderNo
	' sPopUp = Instr( sPopUp,"line 1")
	'If sPopUp = 0  Then
		'	strRowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").RowsCount
		'	strOrderNo= SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebGetCellText( "Order no.",0)
		'
		'		If  strRowCount <> "" Then
		'			AddVerificationInfoToResult  "Order Number" , "Created Successfully"
		'		End If
		'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").ActivateRow 0
		'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebDrillDownColumn "Order no.",0
	'End If


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : OpenExistingOrder_fn
' 	Description : This function creates new order and opens it
'   Created By :  Automation
'	Creation Date :        
'##################################################################################################

Public Function OpenExistingOrder_fn()

    	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	sPopup = "An address change has been processed on this account within the past 7 days"

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "SIS OM Customer Account Portal View","L3" ' clicking on Accounts Screen tab

'	Wait(7)
	Browser("index:=0").Page("index:=0").Sync

	strRowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").RowsCount

	strOrderNo= SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebGetCellText( "Order no.",0)

	If  strRowCount <> "" Then
		AddVerificationInfoToResult  "Order Number" , "Created Successfully"
	End If

	'DataTable.Value("Order_Number") = strOrderNo
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").ActivateRow 0
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebDrillDownColumn "Order no.",0

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : Logout_fn
' 	Description : This function does the logout part
'   Created By :  Automation
'	Creation Date :        
'##################################################################################################

'To Close the siebel browser
Public Function LogoutSiebel_fn()

On Error Resume Next
Err.Clear


'Get Data



	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow


'SiebApplication("Siebel Call Center").SiebMenu("Menu").Exist(0)
SiebApplication("Siebel Call Center").SiebMenu("Menu").Select "File\\File - Logout"

Browser("index:=0").Page("index:=0").Sync

'Close all the browsers after logout
SystemUtil.CloseProcessByName "iexplore.exe"

strSQL = "Select * From Win32_Process Where Name Like '%SiebelAx%'"
 
Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set ProcColl = oWMIService.ExecQuery(strSQL)
		 
For Each oElem in ProcColl
	oElem.Terminate
Next

'	If Err.Number <> 0 Then
'            iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If

End Function



'#################################################################################################
' 	Function Name : SearchAccount_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function SearchAccount_fn()

   Dim adoData	  
'    strSQL = "SELECT * FROM [SearchAccount$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	sAccountNumber = adoData( "Account_Number")  & ""

'	DictionaryTest_G.Item("AccountNo") = "7000318634"
    sAccountNumber = DictionaryTest_G.Item("AccountNo")
'	sAccountName = adoData( "AccountName")  & ""

	If sAccountNumber = "" Then
		AddVerificationInfoToResult "Fail","Account number is null and not retrieved from data base"
		iModuleFail = 1
		Exit Function
	End If

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If	

	
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Query").SiebClick False
    
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Account no.").SiebSetText sAccountNumber
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Go").SiebClick False

	If sAccountName <> "" Then
		sAccountNames = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Account name").SiebGetRoProperty("text"))

			If DictionaryTest_G.Exists("AccountName") Then
				DictionaryTest_G.Item("AccountName")= sAccountNames
			Else
				DictionaryTest_G.add "AccountName",sAccountNames
			End If
	End If

	rowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").RowsCount
	
	If rowCount = 1 Then
		AddVerificationInfoToResult "Search Account - Pass","Account Search is done successfully."
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Search Account - Fail","Account Search is failed"
	End If

	Browser("index:=0").Page("index:=0").Sync

	If   Browser("Siebel Call Center").Dialog("Message from webpage").Exist(4)  Then
			Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click

	End If


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End function


'#################################################################################################
' 	Function Name : SearchAccountSharer_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function SearchAccountSharer_fn()

   Dim adoData	  
    strSQL = "SELECT * FROM [SearchAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sAccountNumber = DictionaryTest_G.Item("AccountNo")

	sAccountName = adoData( "AccountName")  & ""

	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Query").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Account no.").SiebSetText sAccountNumber
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Go").SiebClick False

	If sAccountName <> "" Then
		sAccountNames = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Account name").SiebGetRoProperty("text"))

			If DictionaryTest_G.Exists("AccountName") Then
				DictionaryTest_G.Item("AccountName")= sAccountNames
			Else
				DictionaryTest_G.add "AccountName",sAccountNames
			End If
	End If

	rowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").RowsCount
	
	If rowCount = 1 Then
		AddVerificationInfoToResult "Search Account - Pass","Account Search is done successfully."
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Search Account - Fail","Account Search is failed"
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End function


'#################################################################################################
' 	Function Name : UsedProductServices_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function UsedProductServices_fn()


	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [UsedProductServices$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
'	sSelect = adoData("Select") & ""
	sAction = adoData("Action") & ""
	sEnable= adoData("Enable") & ""
    sEnableAction = adoData("EnableAction") & ""

	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""
	sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
	sLocateColValue = Replace(sLocateColValue,"RootProduct0",DictionaryTest_G.Item("RootProduct0"))
	sLocateColValue = Replace(sLocateColValue,"InstalledId",DictionaryTest_G.Item("ACCNTMSISDN"))
'  sLocateColValue = Replace(sLocateColValue,"InstalledId","447387986719")
	sUIName = adoData( "UIName")  & ""
	sValue = adoData( "Value")  & ""
	sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If


	iIndex = adoData( "Index")  & ""
	If iIndex="" Then
		iIndex=0
	End If

	If DictionaryTest_G.Exists("ActionforUsedProdService") Then
		DictionaryTest_G.Item("ActionforUsedProdService")=sAction
	else
		DictionaryTest_G.add "ActionforUsedProdService",sAction
	End If


	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Used Product/Service")
	'Call ShowMoreButtonUsedProductsServices_fn()
	If sLocateCol <> "" Then

		res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
		If  Cstr(res)<>"True" Then
			 iModuleFail = 1
			AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
			Exit Function
		End If
	End If

	If sAction = "FastOrder" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Used Product/Service").SiebButton("Fast Orders").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List").SiebList("List").ActivateRow 0
			AddVerificationInfoToResult  "Info" ,"Fast Order button is present on right side of Modify button."
			CaptureSnapshot()
			Exit Function
		Else
			AddVerificationInfoToResult  "Error" ,"Fast Order button is not present on right side of Modify button."
			iModuleFail = 1
		End If
	End If

	If sUIName <> "" Then
		If sValue = "ID" Then
			sValue = Trim(Replace(sValue,"ID",DictionaryTest_G.Item("AgreementId")))
		End If
		UpdateSiebList objApplet,sUIName,sValue
	End If
	
	sActEnable =""
	If  sEnableAction<>"" Then '
		 'className = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebTable("Used Product/Service").WebElement("innertext:=" & sEnableAction,"index:=0").WebGetRoProperty ("class")
		 If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebTable("Used Product/Service").WebElement("innertext:=" & sEnableAction,"index:=0","class:=miniBtnUICOff").Exist(5) Then
				sActEnable = "No"
				
				'Call ShowMoreButton_fn()
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List").SiebList("List").ActivateRow 0
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebTable("Used Product/Service").WebElement("innertext:=" & sEnableAction,"index:=0","class:=miniBtnUICOff").highlight
				SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Used Product/Service").SiebButton("micClass:=SiebButton", "repositoryname:=Modify").highlight
'SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List").SiebList("List").ActivateRow 0

				CaptureSnapshot()
		Else
			sActEnable = "Yes"
			SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Used Product/Service").SiebButton("micClass:=SiebButton", "repositoryname:=Modify").highlight
			  CaptureSnapshot()
'			iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , "Class name don't match"
			Exit Function	
		End If
			
		If Ucase(sActEnable) = Ucase(sEnable)  Then
			AddVerificationInfoToResult  "Pass" ,  sEnableAction & " existence is " & sActEnable & " as expected"
			Exit Function
		else
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , sEnableAction & " existence is " & sActEnable & " but expected was " & sEnable
			Exit Function
		End If
	End If


	If(sAction <> "" )Then 
		If(sAction="Modify" or sAction="ModifyBar")Then 
	
			SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Used Product/Service").SiebButton("micClass:=SiebButton", "repositoryname:="&sAction).SiebClick sPopUp
	
			If Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").Exist(3) Then ' checking for PopUp
				Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").WebButton("Confirm").Click
			End If

'			If 	Window("Address Change Alert --").Exist(2) Then
'				Window("Address Change Alert --").Activate
'				Window("Address Change Alert --").WinObject("Internet Explorer_Server").Click 333,164
'			End If
				
		
		Else
	
			SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Used Product/Service").SiebMenu("Menu").SiebSelectMenu sAction
			

				If  Browser("Siebel Call Center_2").Window("CRE Response -- Webpage").Page("CRE Response").WebElement("447389711525: Customer").Exist(5)  Then
					 Browser("Siebel Call Center_2").Window("CRE Response -- Webpage").Page("CRE Response").WebButton("OK").Click
					 AddVerificationInfoToResult  "Info" , "Customer is Upgrade Eligible"
				ElseIf  Browser("Siebel Call Center_2").Window("CRE Response -- Webpage").Page("CRE Response").WebElement("447387960482: Customer").Exist(5) Then
							  iModuleFail = 1
							AddVerificationInfoToResult  "Fail" ,"Customer is Not Upgrade Eligible"
							Exit Function
				End If
	
			If Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").Exist(5) Then ' checking for PopUp
				Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").WebButton("Confirm").Click
			End If
	
		End If
	
		If (sAction="Disconnect") OR (sAction="Post to Pre Migration") OR (sAction="Pre to Post Migration") Then
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Messages").Exist(5) Then
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Messages").SiebButton("Accept").SiebClick False
			End If
		End If
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : VerifyAddressLine_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function VerifyAddressLine_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sAddressLine1= Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Address Line 1").GetRoProperty("text"))

	If DictionaryTest_G.Item("AddressLine") = sAddressLine1 Then
		AddVerificationInfoToResult "Pass","Address line 1 " & sAddressLine1 & " is displayed correctly in Account summary page"
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Fail","Address line 1 " & sAddressLine1 & " is not displayed correctly in Account summary page"
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : AccountHierarchyList_fn
' 	Description : This function is used to select list of items from Account Summary page and then perform the action
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AccountHierarchyList_fn()

Dim adoData	  
Dim strSQL
Dim sAccountNumber

    strSQL = "SELECT * FROM [AccountHierarchy$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sAccountNumber = adoData( "AccountNumber")  & ""

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Hierarchy Analysis List View","L3"
	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Hierarchy").SiebApplet("Account Org Hierarchy").SiebButton("Add").SiebClick False ' Clicking on Add Button

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Hierarchy").SiebApplet("Pick Account").SiebButton("Query").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Hierarchy").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").SiebSetText sAccountNumber
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Hierarchy").SiebApplet("Pick Account").SiebButton("Go").SiebClick False

	rwNum = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Hierarchy").SiebApplet("Account Org Hierarchy").SiebList("List").SiebListRowCount

	For  loopVar = 0 to rwNum - 1
	
	strParentAccount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Hierarchy").SiebApplet("Account Org Hierarchy").SiebList("List").SiebGetCellText("Parent Account",loopVar)
	
	If Trim(strParentAccount) = DictionaryTest_G.Item("AccountName") Then
		AddVerificationInfoToResult "Pass","Child is added successfully under Parent : " & strParentAccount
		Exit For
	End If
	
Next

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


''#################################################################################################
' 	Function Name : GoToProfiles_fn
' 	Description : This function is used to select list of items from Account Summary page and then perform the action
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function GoToProfiles_fn()

'Dim adoData	  
'Dim strSQL
'Dim sAccountNumber

'    strSQL = "SELECT * FROM [DirectDebit$] WHERE  [RowNo]=" & iRowNo
''	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow


	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").ActivateRow 0


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

''#################################################################################################
' 	Function Name : DirectDebit_fn
' 	Description : This function is used to select list of items from Account Summary page and then perform the action
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function DirectDebit_fn()

Dim adoData	  
Dim strSQL
Dim sAccountNumber

    strSQL = "SELECT * FROM [DirectDebit$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sPaymentMethod = Trim(adoData( "PaymentMethod")  & "")
	sSortCode = Trim(adoData( "SortCode")  & "")
	sAccountNumber = Trim(adoData( "AccountNumber")  & "")
	sAccountName = Trim(adoData( "AccountName")  & "")
	sBranch = trim(adoData( "Branch")  & "")
	sBankName = Trim(adoData( "BankName")  & "")
	sReEstablishment = Trim(adoData( "ReEstablishment")  & "")
	sValidateButton = Trim(adoData( "ValidateButton")  & "")
	sMandateID = Trim(adoData( "MandateID")  & "")
	sMandateStatus = Trim(adoData( "MandateStatus")  & "")
	sNavigate = Trim(adoData( "Navigate")  & "")
	sReadOnly = adoData( "ReadOnly")  & ""


	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").ActivateRow 0



	If Ucase(sReEstablishment) <> "YES" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment method").SiebSelect sPaymentMethod
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebButton("Update Bank Details").SiebClick False
				Browser("index:=0").Page("index:=0").Sync
				If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Siebel").Exist(5) Then
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Siebel").SiebText("Bank name").SiebSetText sBankName
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Siebel").SiebText("Branch").SiebSetText sBranch
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Siebel").SiebText("Sort code").SiebSetText sSortCode
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Siebel").SiebText("Account name").SiebSetText sAccountName
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Siebel").SiebText("Account no.").SiebSetText sAccountNumber
					AddVerificationInfoToResult  "Info" ,"Bank Details are filled successfully."
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Siebel").SiebButton("OK").SiebClick False
				Else
					AddVerificationInfoToResult  "Error" ,"Update Bank Details button is not clicked successfully."
					iModuleFail = 1
					Exit Function
				End If

		Else
					strAccountStatus = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Account status").GetROProperty("text")
					If strAccountStatus<>"Valid" Then
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebButton("Validate Bank Details").SiebClick False
							strAccountStatus = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Account status").GetROProperty("text")
							AddVerificationInfoToResult  "Info" ,"Validate Bank Details Status is." &strAccountStatus
					End If

						
	''validate Bank Name					
						strBankName = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Bank name").GetROProperty("text")
						'If strBankName <> "" Then
							If sBankName<>"" Then
								If instr(sBankName,strBankName)>0 Then
										AddVerificationInfoToResult  "Pass" ,"Bank Name is " &strBankName
										else
											AddVerificationInfoToResult  "Fail" ,"Bank Name is " &strBankName
											iModuleFail = 1
											Exit Function

								End If
'						else
'							AddVerificationInfoToResult  "Pass" ,"Bank Name is " &strBankName
							End if
'						Else
'							AddVerificationInfoToResult  "Error" ,"Bank Name is null"
'							iModuleFail = 1
'							Exit Function
'						End If
						

		''validate Bank Branch	  sBranch			
						strBranch = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Branch").GetROProperty("text")
					'	If strBranch <> "" Then
								If sBranch<>"" Then
									If instr(sBranch,strBranch)>0 Then
											AddVerificationInfoToResult  "Pass" ,"Bank Branch is " &strBranch
											else
											AddVerificationInfoToResult  "Fail" ,"Bank Branch is " &strBranch
												iModuleFail = 1
												Exit Function
									End If
'								else
'								AddVerificationInfoToResult  "Pass" ,"Bank Branch is " &strBranch
								End If
'						Else
'							AddVerificationInfoToResult  "Error" ," Bank Branch is null"
'							iModuleFail = 1
'							Exit Function
'						End If
						

	'validate sort code					
						strSortCode = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Sort code").GetROProperty("text")
						If strSortCode <> "" Then
							If sSortCode<>"" Then
									If instr(sSortCode,strSortCode)>0 Then
											AddVerificationInfoToResult  "Pass" ,"SortCode is " &strSortCode
											else
											AddVerificationInfoToResult  "Fail" ,"SortCode is " &strSortCode
												iModuleFail = 1
												Exit Function
									End If
								else
								AddVerificationInfoToResult  "Pass" ,"SortCode is " &strSortCode
								End If
							
						Else
							AddVerificationInfoToResult  "Error" ," SortCode is null"
							iModuleFail = 1
							Exit Function
						End If
						

		'validate Account Name				
						strAccountName = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Account name").GetROProperty("text")
						If strAccountName <> "" Then
							If sAccountName<>"" Then
									If instr(sAccountName,strAccountName)>0 Then
											AddVerificationInfoToResult  "Pass" ,"Account Name is " &strAccountName
											else
											AddVerificationInfoToResult  "Fail" ,"Account Name is " &strAccountName
												iModuleFail = 1
												Exit Function
									End If
								else
								AddVerificationInfoToResult  "Pass" ,"Account Name is " &strAccountName
								End If
							
						Else
							AddVerificationInfoToResult  "Error" ," Account Name is null"
							iModuleFail = 1
							Exit Function
						End If

		'	validate Account Num			
						strAccountNum = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Account no.").GetROProperty("text")
					If strAccountNum <> "" Then
							If sAccountNumber<>"" Then
									If instr(sAccountNumber,strAccountNum)>0 Then
											AddVerificationInfoToResult  "Pass" ,"Account Num is " &strAccountNum
											else
											AddVerificationInfoToResult  "Fail" ,"Account Num is " &strAccountNum
												iModuleFail = 1
												Exit Function
									End If
								else
								AddVerificationInfoToResult  "Pass" ,"Account Num is " &strAccountNum
								End If
							
						Else
							AddVerificationInfoToResult  "Error" ," Account Num is null"
							iModuleFail = 1
							Exit Function
						End If
						

		'validate Mandate Status	  sMandateStatus			
					strMandateStatus =	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Mandate status").GetROProperty("text")
					If strMandateStatus <> "" Then
							If sMandateStatus<>"" Then
									If instr(sMandateStatus,strMandateStatus)>0 Then
											AddVerificationInfoToResult  "Pass" ,"MandateStatus is " &strMandateStatus
											else
											AddVerificationInfoToResult  "Fail" ,"MandateStatus is " &strMandateStatus
												iModuleFail = 1
												Exit Function
									End If
								else
								AddVerificationInfoToResult  "Pass" ,"MandateStatus is " &strMandateStatus
								End If
							
						Else
							AddVerificationInfoToResult  "Error" ," MandateStatus is null"
							iModuleFail = 1
							Exit Function
						End If

			'validate Mandate ID		
						strMandateID = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Mandate ID").GetROProperty("text")
						If strMandateID <> "" Then
							strMandateIdNew = Split(strMandateID,"-")(1)
								If  sMandateID <>""Then
										if (strMandateIdNew=sMandateID) Then
											AddVerificationInfoToResult  "Pass" ," Mandate ID is as expected : "&sMandateID
										Else
											AddVerificationInfoToResult  "Fail" ," Mandate ID expected is  : "&sMandateID &"Actual is  : "&strMandateIdNew
										End If
								End If
						

						Else
							AddVerificationInfoToResult  "Error" ," MandateID is null"
							iModuleFail = 1
							Exit Function
						End If

		End if

	Browser("index:=0").Page("index:=0").Sync

	If (sReadOnly = "Yes") Then
				'for Account name
					strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Account name").getRoProperty("isenabled")
                    		If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Account name is not  editable.. as Expected."
								CaptureSnapshot()
							Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Account name is editable."
								CaptureSnapshot()
								Exit Function

						End If
				'for Account number
					strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Account no.").getRoProperty("isenabled")
                		If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Account no. is not  editable.. as Expected."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Account no. is editable."
								CaptureSnapshot()
								Exit Function

						End If	
					'for Account status
					strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Account status").getRoProperty("isenabled")
                    	If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Account status is not  editable.. as Expected."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Account status is editable."
								CaptureSnapshot()
								Exit Function

						End If	
					'for Bank name
					strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Bank name").getRoProperty("isenabled")
						 If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Bank name is not  editable.. as Expected."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Bank name is editable."
								CaptureSnapshot()
								Exit Function

						End If	
					'for Branch 
					strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Branch").getRoProperty("isenabled")
						If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Branch is not  editable.. as Expected."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Branch is editable."
								CaptureSnapshot()
								Exit Function

						End If
					'for Sort Code
					strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Sort code").getRoProperty("isenabled")
                    	If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Sort code is not  editable.. as Expected."
								CaptureSnapshot()
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Sort code is editable."
								CaptureSnapshot()
								Exit Function

						End If

	End If

	If Ucase(sValidateButton) = "YES" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebButton("Validate Bank Details").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
	End If
	


			If Ucase(sNavigate) = "YES" Then
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "SIS OM Customer Account Portal View","L3"
			End If


	CaptureSnapshot()
'err.clear
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


''#################################################################################################
' 	Function Name : CheckMandateStatus_fn
' 	Description : This function is used to select list of items from Account Summary page and then perform the action
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function CheckMandateStatus_fn()

Dim adoData	  
Dim strSQL
Dim sAccountNumber

    strSQL = "SELECT * FROM [DirectDebit$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sMandateID = Trim(adoData( "MandateID")  & "")
	sMandateStatus = Trim(adoData( "MandateStatus")  & "")
'	sNavigate = Trim(adoData( "Navigate")  & "")
'	sReadOnly = adoData( "ReadOnly")  & ""


	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").ActivateRow 0


		'validate Mandate Status	  sMandateStatus			
					strMandateStatus =	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Mandate status").GetROProperty("text")
					If strMandateStatus <> "" Then
							If sMandateStatus<>"" Then
									If instr(sMandateStatus,strMandateStatus)>0 Then
											AddVerificationInfoToResult  "Pass" ,"MandateStatus is " &strMandateStatus
										else
											AddVerificationInfoToResult  "Fail" ,"MandateStatus is " &strMandateStatus
												iModuleFail = 1
												Exit Function
									End If
								else
								AddVerificationInfoToResult  "Pass" ,"MandateStatus is " &strMandateStatus
								End If
							
						Else
							AddVerificationInfoToResult  "Error" ," MandateStatus is null"
							iModuleFail = 1
							Exit Function
						End If

			'validate Mandate ID		
						strMandateID = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Mandate ID").GetROProperty("text")
						If strMandateID <> "" Then
							strMandateIdNew = Split(strMandateID,"-")(1)
								If  sMandateID <>""Then
										if (strMandateIdNew=sMandateID) Then
											AddVerificationInfoToResult  "Pass" ," Mandate ID is as expected : "&sMandateID
													If DictionaryTest_G.Exists("MandateID") Then
															DictionaryTest_G.Item("MandateID")=strMandateID 
														else
															DictionaryTest_G.add "MandateID",strMandateID 
													End If
										Else
											AddVerificationInfoToResult  "Fail" ," Mandate ID expected is  : "&sMandateID &"Actual is  : "&strMandateIdNew
										End If
								End If
						

						Else
							AddVerificationInfoToResult  "Error" ," MandateID is null"
							iModuleFail = 1
							Exit Function
						End If

	CaptureSnapshot()
'err.clear
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



''#################################################################################################
' 	Function Name : CaptureDirecDebitFields
' 	Description : This function is used to select list of items from Account Summary page and then perform the action
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function CaptureDirecDebitFields()

Dim adoData	  
Dim strSQL
Dim sAccountNumber

    strSQL = "SELECT * FROM [DirectDebit$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sMandateID = Trim(adoData( "MandateID")  & "")
	sMandateStatus = Trim(adoData( "MandateStatus")  & "")
'	sNavigate = Trim(adoData( "Navigate")  & "")
'	sReadOnly = adoData( "ReadOnly")  & ""


	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").ActivateRow 0


		'validate Mandate Status	  sMandateStatus			
					strMandateStatus =	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Mandate status").GetROProperty("text")
					If strMandateStatus <> "" Then
							If sMandateStatus<>"" Then
									If instr(sMandateStatus,strMandateStatus)>0 Then
											AddVerificationInfoToResult  "Pass" ,"MandateStatus is " &strMandateStatus
											else
											AddVerificationInfoToResult  "Fail" ,"MandateStatus is " &strMandateStatus
												iModuleFail = 1
												Exit Function
									End If
								else
								AddVerificationInfoToResult  "Pass" ,"MandateStatus is " &strMandateStatus
								End If
							
						Else
							AddVerificationInfoToResult  "Error" ," MandateStatus is null"
							iModuleFail = 1
							Exit Function
						End If

			'validate Mandate ID		
						strMandateID = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebText("Mandate ID").GetROProperty("text")
						If strMandateID <> "" Then
							strMandateIdNew = Split(strMandateID,"-")(1)
								If  sMandateID <>""Then
										if (strMandateIdNew=sMandateID) Then
											AddVerificationInfoToResult  "Pass" ," Mandate ID is as expected : "&sMandateID
										Else
											AddVerificationInfoToResult  "Fail" ," Mandate ID expected is  : "&sMandateID &"Actual is  : "&strMandateIdNew
										End If
								End If
						

						Else
							AddVerificationInfoToResult  "Error" ," MandateID is null"
							iModuleFail = 1
							Exit Function
						End If

	CaptureSnapshot()
'err.clear
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : CancelAllOPenAndPendingOrders_fn
' 	Description : This function is used to select list of items from Account Summary page and then perform the action
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CancelAllOPenAndPendingOrders_fn()

Dim adoData	  
Dim strSQL
Dim sAccountNumber

    strSQL = "SELECT * FROM [DirectDebit$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow


	Browser("index:=0").Page("index:=0").Sync
	If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("_svf3").Image("Show more").Exist(2)  Then
		 Browser("Siebel Call Center").Page("Siebel Call Center").Frame("_svf3").Image("Show more").Click
	End If

	
	rwNum =SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebListRowCount

	For  loopVar = 0 to rwNum - 1
	
	strStatus=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebGetCellText("Status",loopVar)
	
	If strStatus="Open" or strStatus="Pending" Then
		sOrderNo=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Orders").SiebList("List").SiebGetCellText("Order no.",loopVar)
		If DictionaryTest_G.Exists("OrderNo") Then
			DictionaryTest_G.Item("OrderNo")=sOrderNo
		else
			DictionaryTest_G.add "OrderNo",sOrderNo
		End If
	End If
	Call ExecuteDBQueryOrderCancel_fn()
	
	Next

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : ModifyDetails_fn
' 	Description : This function is used to modify several details on Account page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ModifyDetails_fn()

Dim adoData	  
Dim strSQL
Dim sLastName
Dim sModifyType
Dim sLastNameApp
Dim sLastNameApp1
Dim sEmailApp
Dim sEmailApp1
Dim sEmail

    strSQL = "SELECT * FROM [ModifyDetails$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sModifyType = Trim(adoData( "ModifyType")  & "")
	sLastName = Trim(adoData( "LastName")  & "")
	sPostCode = Trim(adoData( "PostCode")  & "")
	sEmail = Trim(adoData( "Email")  & "")

sPopUp = adoData( "Popup")  & ""


	If Ucase(sPopUp)="FALSE" or sPopUp=""Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If


	'Flow

	Browser("index:=0").Page("index:=0").Sync


	If sModifyType = "LastName" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
			sLastNameApp = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("Last name").GetROProperty("text") ''Fetching Last NAme
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("Last name").SiebSetText sLastName '' Setting last name
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebButton("Save").SiebClick False
		
			Browser("index:=0").Page("index:=0").Sync
		
			sLastNameApp1 = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("Last name").GetROProperty("text")'Fetching LAst name after changing it
		
			If sLastNameApp <> sLastNameApp1 Then
				AddVerificationInfoToResult "Pass","Last Name is changed successfully from " & sLastNameApp & " to " & sLastNameApp1
			Else
				AddVerificationInfoToResult "Fail","Last Name is not changed successfully."
					iModuleFail = 1
			End If

			If Err.Number <> 0 Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , Err.Description
			End If
	End If

	If sModifyType = "Email" Then

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
			sEmailApp = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("Email").GetROProperty("text") ''Fetching Email
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebButton("Update Email").SiebClick False
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Siebel").SiebText("Email").SiebSetText "changedemail"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Siebel").SiebPicklist("SiebPicklist").SetText "ymail.com"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Siebel").SiebText("Re-type Email").SiebSetText "changedemail"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Siebel").SiebPicklist("SiebPicklist_2").SetText "ymail.com"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Siebel").SiebButton("OK").SiebClick False


			Browser("index:=0").Page("index:=0").Sync
		
			sEmailApp1 = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("Email").GetROProperty("text")'Fetching Email after changing it
		
			If sEmailApp <> sEmailApp1 Then
				AddVerificationInfoToResult "PAss","Email is changed successfully from " & sEmailApp & " to " & sEmailApp1
			Else
				AddVerificationInfoToResult "Fail","Email is not changed successfully."
					iModuleFail = 1
			End If

			If Err.Number <> 0 Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , Err.Description
			End If
	End If



	If sModifyType = "Address" Then
		On error Resume Next
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Business Address View","L3"
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").SiebButton("New").SiebClick False
		
		'SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Add Address").SiebText("SiebText").SiebSetText sPostCode
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SiebSetText sPostCode 
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").ProcessKey "EnterKey"
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Add Address").SiebButton("Go").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Add Address").SiebButton("OK").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").SiebList("List").SiebCheckbox("Primary").SiebSetCheckbox "On"
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").SiebList("List").SiebCheckbox("Primary").SetOn
		err.clear

		If Browser("Siebel Call Center").Dialog("Message from webpage").Static("Customer's address has").Exist(3) Then
			AddVerificationInfoToResult  "Info" , "Customer address is changed successfully."
			CaptureSnapshot()
		Else
			AddVerificationInfoToResult  "Fail", "Customer address is not changed successfully."
		End If

	End If

	CaptureSnapshot()
		

    
End Function

'#################################################################################################
' 	Function Name : BillingProfileVerify_fn
' 	Description : This function is used to validate Payment menthod type in billing profile
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BillingProfileVerify_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewBillingProfile$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
'	sPrePostPay = adoData( "PrePostPay")  & ""
'	sPaymentMethod = adoData( "PaymentMethod")  & ""

	'Flow

	'*****************************************************************************
	'Check billing profile Payment Method

'	If sPrePostPay = "Prepaid" Then
'		sPaymentMethodVal = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebList("List").SiebPicklist("Payment method").SiebGetRoProperty("activeitem"))
'		If sPaymentMethodVal = sPaymentMethod Then
'			AddVerificationInfoToResult "Pass","Payment method " & sPaymentMethodVal & " is updated correctly against " & sPrePostPay
'		Else
'			AddVerificationInfoToResult "Fail","Payment method is not updated correctly"
'			 iModuleFail = 1
'		End If
'	ElseIf sPrePostPay = "Postpaid" Then
'		sPaymentMethodVal = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles").SiebList("List").SiebPicklist("Payment method").SiebGetRoProperty("activeitem"))
'		If sPaymentMethodVal = sPaymentMethod Then
'			AddVerificationInfoToResult "Pass","Payment method " & sPaymentMethodVal & " is updated correctly against " & sPrePostPay
'		Else
'			AddVerificationInfoToResult "Fail","Payment method is not updated correctly"
'			 iModuleFail = 1
'		End If
'	End If


	''''''++++++++++++++++++++++++++++++NewCode+++++++++++++++++++++++++++++++++++++++

		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing Profiles")
			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

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



	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : CreateCustComms_fn
' 	Description : This function is used to used to create new customer comms
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CreateCustComms_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewCustComms$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sCategory = adoData( "Category")  & ""
	sSubCategory = adoData( "Subcategory")  & "" 
	sResolution = adoData( "Resolution")  & ""
	sDPAValidation = adoData( "DPA_validation")  & ""
	sType = adoData( "Type")  & ""
	sClickAccount = adoData( "ClickAccount")  & ""
	sCheckEligibility = adoData( "CheckEligibility")  & ""
	sDPAValidationInLineItems = adoData( "DPAValidationInLineItems")  & ""
	sVerifyStatus = adoData( "VerifyStatus")  & ""
	sSendOTACSMS = adoData( "SendOTACSMS")  & ""
	sEnabledOTACSMS = adoData( "EnabledOTACSMS")  & ""

	    sPopUp = adoData( "Popup")  & ""    
        If Ucase(sPopUp)="NO" Then
            sPopUp="FALSE"
            sPopUp=Cbool(sPopUp)
        End If

	MN =Left( MonthName(Month(Date)),3)
	DateRequested = Date					
	DateRequested = day(DateRequested) + 2 & "/" & MN & "/" &  year(DateRequested)				
	sAppointmentDate = DateRequested
	
	sInstalledID =DictionaryTest_G("ACCNTMSISDN") 


	'Flow

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List").SiebButton("New").SiebClick False

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List").SiebList("List").SiebDrillDownColumn "ID",0

		If sDPAValidationInLineItems="Passed" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("DPA validation").SiebSelect sDPAValidationInLineItems
				CaptureSnapshot()
		End If

		If sInstalledID<>"" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebText("Installed ID").SiebSetText sInstalledID
		End If	

		If sCategory<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Category").SiebSelect sCategory
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Sub-category").SiebSelect sSubCategory
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Resolution").SiebSelect sResolution			
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebCalendar("Appointment Date").SiebSetText sAppointmentDate
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Appointment Slot").SiebSelect "AM"
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebTextArea("Comments").SetText "Solved"
                       
			If sInstalledID<>"" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebText("Installed ID").SiebSetText sInstalledID
			End If		
			If sDPAValidation <> "" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("DPA validation").SiebSelect sDPAValidation
			End If

					If sVerifyStatus <> "" Then
						 	If sVerifyStatus = "Closed" Then
							'SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebButton("Close").SiebClick False
							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebButton("Close").SiebClick False
							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebButton("Close").SiebClick False

							wait 1
							SiebApplication("Siebel Call Center").SiebToolbar("HIQuery").Click "ExecuteQuery" ' clicking on Refresh Button							
							End If
						
							sActStatus = Trim(SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Status").GetROProperty("activeitem"))
							If Ucase(sVerifyStatus) = Ucase(sActStatus)  Then
											AddVerificationInfoToResult "Pass","Customer Comms Status is " &sActStatus& " as expected"
							Else
											AddVerificationInfoToResult "Fail","Customer Comms Status is " &sActStatus& " not as expected"  
											 iModuleFail = 1
							End If
					End If

					If sEnabledOTACSMS = "Y" Then
							If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebButton("Send OTAC SMS").isenabled = True then
							AddVerificationInfoToResult "Pass","Send OTA SMS button is enabled when status is In-progress"
							Else
										AddVerificationInfoToResult "Fail","Send OTA SMS button is not enabled when status is In-progress"
										 iModuleFail = 1
							 End If
						
					End If

					If sSendOTACSMS = "Y" Then
						SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebButton("Send OTAC SMS").SiebClick False
						SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Send OTAC").SiebButton("OK").SiebClick sPopUp
						SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Accounts").SiebList("List").DrillDownColumn "Name",0
						Browser("index:=0").Page("index:=0").Sync
					End If
			

		End If

		If sCheckEligibility = "Y" Then
			  SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebButton("Check Eligibility").SiebClick sPopUp
		End If
		
				If sClickAccount = "Y" Then
					Set objApplet =SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Accounts")
					objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0
					SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Accounts").SiebList("List").DrillDownColumn "Name",0 
					Browser("index:=0").Page("index:=0").Sync
				End If
			
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : ChangePaymentMethod_fn
' 	Description : This function is used to used to change the payment method in Profiles tab.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ChangePaymentMethod_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ChangePaymentMethod$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sPaymentMethod = adoData( "PaymentMethod")  & ""	
	sExpectedPaymentMethod =  adoData( "ExpectedPaymentMethod")  & ""	

		sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

	'Flow
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3" ' clicking on Profiles tab
		Browser("index:=0").Page("index:=0").Sync
		RowC = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebListRowCount

				For i = 0 to RowC-1
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").ActivateRow i
					Browser("index:=0").Page("index:=0").Sync
						If sPaymentMethod <> "" and sExpectedPaymentMethod <>"" Then
							sGetPaymentMethod  =Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment method").GetROProperty("activeitem"))
							CaptureSnapshot()
								If sGetPaymentMethod = sExpectedPaymentMethod Then
									SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment method").SiebSelect sPaymentMethod
									sNewPaymentMethod  =Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment method").GetROProperty("activeitem"))
									Exit For
								End If
										If Trim(sPaymentMethod) = Trim(sNewPaymentMethod)  Then
											AddVerificationInfoToResult "Info","Payment Method is changed from " &sGetPaymentMethod & " to " &sNewPaymentMethod
													If sPaymentMethod = "Credit/Debit Card"  Then
															SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebButton("New").SiebClick sPopUp

													End If
										Else
											AddVerificationInfoToResult "Fail","Payment Method is not changed from " &sGetPaymentMethod & " to " &sNewPaymentMethod
											 iModuleFail = 1
										End If
									CaptureSnapshot()	
						ElseIf sBillingProfileStatus <>"" and sExpectedBillingProfileStatus <>"" Then
										sGetBillingProfileStatus =Trim (SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Billing profile status").GetROProperty("activeitem"))
										CaptureSnapshot()
											If sGetBillingProfileStatus = sExpectedBillingProfileStatus Then
												SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Billing profile status").SiebSelect sBillingProfileStatus
												sNewBillingProfileStatus  =Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Billing profile status").GetROProperty("activeitem"))
												CaptureSnapshot()
												Exit For
											End If
													If Trim(sBillingProfileStatus) = Trim(sNewBillingProfileStatus)  Then
														AddVerificationInfoToResult "Info","Billing Profile Status is changed from " &sGetBillingProfileStatus & " to " &sNewBillingProfileStatus
													Else
														AddVerificationInfoToResult "Fail","Billing Profile Status is not changed from " &sGetBillingProfileStatus & " to " &sNewBillingProfileStatus
														 iModuleFail = 1
													End If
												CaptureSnapshot()	

						End If
					Next

		
  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : UpdateContactPhoneNumber_fn
' 	Description : This function is used to used to insert the moble phone number in contacts tab on Accounts page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function UpdateContactPhoneNumber_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ContactPhoneNumber$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sMobileNumber = adoData( "MobileNumber")  & ""	


	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").SiebText("Mobile phone no.").SiebSetText sMobileNumber
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Save").Click

	If Browser("Siebel Call Center").Dialog("Siebel").Exist(3) Then
		Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").SiebText("Mobile phone no.").SiebSetText sMobileNumber
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Save").Click
	End If

	If Browser("Siebel Call Center").Dialog("Siebel").Static("The selected record has").Exist(3) Then
		Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
	End If

	AddVerificationInfoToResult "Info","Mobile Number " &sMobileNumber & " is updated successfully. " 

		CaptureSnapshot()			
  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : ProductServicesVerify_fn()
' 	Description : This function is used to validate all the products on Product/Services tab in Accounts page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function ProductServicesVerify_fn()

	Dim rwNum
	Dim loopVar
	Dim strProduct
	Dim sProduct
	Dim adoData
	Dim strSQL

		  
    strSQL = "SELECT * FROM [ProductServices$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "SIS OM Products & Services View (Service)","L3" ' NAvigate to Product/Services tab

	

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Products and Services").SiebList("List").Exist(2) Then
		AddVerificationInfoToResult "Info","Product/Services list is displayed successfully."
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Fail","Product/Services list is not displayed after clicking on Product/Services tab on Accounts page."
		Exit Function
	End If
	


	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Products and Services").SiebButton("ToggleListRowCount").Click ' clicking on Toggle button to expand
	'Call ShowMoreButton_fn()

	Browser("index:=0").Page("index:=0").Sync

	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Products and Services")

	Do while Not adoData.Eof
		sLocateColExpand = adoData( "LocateColExpand")  & ""
		sLocateColExpandValue = adoData( "LocateColExpandValue")  & ""
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sLocateColValue = Replace(sLocateColValue,"RootProduct0",DictionaryTest_G.Item("ROOT_PRODUCT"))
		sLocateColValue = Replace(sLocateColValue,"InstalledId",DictionaryTest_G.Item("ACCNTMSISDN"))
'		sExpand = adoData( "Expand")  & ""		
		sCollapse = adoData( "Collapse")  & ""
		sExist = adoData( "Exist")  & ""
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
		iIndex = 0
		If sLocateColExpand <> "" Then
			'sLocateColExpand = Replace(sLocateColExpand,sLocateColExpand,DictionaryTest_G.Item("ROOT_PRODUCT"))
			res=LocateColumns (objApplet,sLocateColExpand,sLocateColExpandValue,iIndex)
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Products and Services").ExpandRow
		End If
		If sLocateCol <> "" Then
		
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			
			If sUIName <> "" Then
				UpdateSiebList objApplet,sUIName,sValue
				If DictionaryTest_G.Exists("AgreementId") Then
					DictionaryTest_G.Item("AgreementId")=sValue
				else
					DictionaryTest_G.add "AgreementId",sValue
				End If
				AddVerificationInfoToResult  "Info" ,"Agreement Id value captured from ProductServices view is : " & sValue
				CaptureSnapshot()
			End If

			If sExist = "" Then
				If  res=True Then
					 AddVerificationInfoToResult  "Pass" , sLocateCol & " - " & sLocateColValue & " found in the list as expected"
				else
					 iModuleFail = 1
					AddVerificationInfoToResult  "Error" , sLocateCol & " - " & sLocateColValue & " not found in the list."				
				End If 
			else
				If sExist = "False" Then
					sExist = "False-Row Not Exist"
				End If
				If Ucase(cstr(res)) = Ucase (sExist) Then
					AddVerificationInfoToResult  "Pass" , sLocateCol & " - " & sLocateColValue & " existence is " & sExist & " as expected"
				else
					 iModuleFail = 1
					AddVerificationInfoToResult  "Pass" , sLocateCol & " - " & sLocateColValue & " existence is " & cstr(res) & " but expected is " & sExist
				End If
			End If
		End If


		If sCollapse="Y" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Products and Services").SiebMenu("Menu").SiebSelectMenu "Run Query"
'			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectCollapseRow sCollapse,iIndex
		End If
		adoData.MoveNext
	Loop

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : UpdateEmail_fn
' 	Description : This function is used to used to update email id on Accounts page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function UpdateEmail_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [ContactPhoneNumber$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	sMobileNumber = adoData( "MobileNumber")  & ""	

	'Flow
	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Update Email").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebText("Email").SiebSetText "XYZ"
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebPicklist("SiebPicklist").SiebSelect "hotmail.co.uk"
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebText("Re-type Email").SiebSetText "XYZ"
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebPicklist("SiebPicklist_2").SiebSelect "hotmail.co.uk"
	CaptureSnapshot()
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebButton("OK").SiebClick False

	AddVerificationInfoToResult "Info","Email Id is updated successfully." 

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : AddressEditable_fn
' 	Description : This function is used to used to check whether  address can be editable for Anonymous customer under differenrt conditions
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AddressEditable_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AddressUpdate$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sEditable = adoData( "Editable")  & ""	
'    sNonAnonymous = adoData( "NonAnonymous")  & ""
	sPost_Code = adoData( "Post_Code")  & ""
	sLastName = adoData( "LastName")  & ""
	sPopUp = adoData( "PopUp")  & ""

		If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If

	'Flow
	Browser("index:=0").Page("index:=0").Sync     'For Account Summary

	If sEditable = "Address Line 1" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup
		If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebElement("innertext:=validate","index:=0","class:=miniBtnUICOff","html tag:=SPAN").Exist(5) Then
			AddVerificationInfoToResult "Info","Validate button is disabled for Anonymous customer at Address Line 1 under Account Address window"
			CaptureSnapshot()
		Else
			AddVerificationInfoToResult "Fail","Validate button is enabled for Anonymous customer at Address Line 1 under Account Address window."
			iModuleFail = 1
		End If
'			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Addresses").SiebButton("OK").SiebClick False
	End If

If sEditable = "Address Line 1 NonAnonymous" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SetText sPost_Code
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").ProcessKey "EnterKey"
		Browser("index:=0").Page("index:=0").Sync
        SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Addresses").SiebButton("Add").SiebClick False
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Addresses").SiebButton("OK").SiebClick False
End If

	If sEditable = "Addresses" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Business Address View","L3"
		If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebElement("Validate").Exist(5) Then
			AddVerificationInfoToResult "Info","Validate button is disabled for Anonymous customer at Addresses tab."
			CaptureSnapshot()
		Else
			AddVerificationInfoToResult "Fail","Validate button is enabled for Anonymous customer at Addresses tab."
			iModuleFail = 1
		End If
	End If

	If sEditable = "Contacts NonAnonymous" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
		CaptureSnapshot()
		Browser("index:=0").Page("index:=0").Sync
    	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Add").SiebClick False
		CaptureSnapshot()
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").SiebText("Last name").SiebSetText sLastName
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").SiebText("Last name").ProcessKey "EnterKey"
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebButton("OK").SiebClick sPopUp
		Browser("index:=0").Page("index:=0").Sync
	End If
    
	If sEditable = "Age and Id Verification" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Age Verification View","L3"

		If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("innertext:=validate","index:=0","class:=miniBtnUICOff","html tag:=SPAN").Exist(5) Then
			AddVerificationInfoToResult "Info","Validate button is disabled for Anonymous customer at Age and ID Verification tab under Account addresses."
			CaptureSnapshot()
		Else
			AddVerificationInfoToResult "Fail","Validate button is enabled for Anonymous customer at Age and ID Verification tab under Account addresses."
			iModuleFail = 1
		End If
	End If

	If sEditable = "Age and Id Verification NonAnonymous" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Age Verification View","L3"
			CaptureSnapshot()
			Browser("index:=0").Page("index:=0").Sync
'			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Addresses").SiebClick False
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Addresses").SiebButton("New").isenabled = True then
						CaptureSnapshot()
						AddVerificationInfoToResult "Pass","New Button is editable.."
			Else
						iModuleFail = 1
						AddVerificationInfoToResult "Fail","New Button is not editable."
						Exit Function
			End If
	End If


	If sEditable = "Credit Vetting" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Credit Vetting View","L3"

		If Browser("Siebel Call Center").Dialog("Message from webpage").Static("Please enter a minimum").Exist(5) Then
			Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
		End If

		If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("innertext:=validate","index:=0","class:=miniBtnUICOff","html tag:=SPAN").Exist(5) Then
			AddVerificationInfoToResult "Info","Validate button is disabled for Anonymous customer at Credit Vetting  tab under Address History."
			CaptureSnapshot()
		Else
			AddVerificationInfoToResult "Fail","Validate button is enabled for Anonymous customer at Credit Vetting  tab under Address History."
			iModuleFail = 1
		End If
	
	End If

	If sEditable = "CreditVetting_NonAnonymous" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Credit Vetting View","L3"
			CaptureSnapshot()
			Browser("index:=0").Page("index:=0").Sync
			If Browser("Siebel Call Center_2").Dialog("Message from webpage").Static("IF A CREDIT VET IS REQUIRED").Exist then
						Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Click
			End If
			if SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Address History").SiebButton("New").isenabled = True then
							CaptureSnapshot()
							AddVerificationInfoToResult "Pass","New Button is editable.."
			Else
							iModuleFail = 1
							AddVerificationInfoToResult "Fail","New Button is not editable."
							Exit Function
			End if
		End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : AccountBillingProfileAboutRecord_fn
' 	Description : This function is used to used to click on Profile tab on Account page,capture Payment method  and select About Record after selecting billing profile.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AccountBillingProfileAboutRecord_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sProfilesPaymentMethodStatusCheck = adoData( "StatusCheck")  & ""	

	If sProfilesPaymentMethodStatusCheck <> "" Then ' After performing Tokenization
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3" ' clicking on Profile tab
'		SiebApplication("Siebel Call Center").SiebToolbar("HIQuery").Click "ExecuteQuery"
		Browser("index:=0").Page("index:=0").Sync
	
		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").Exist(3) Then
			sPaymentMethod1 = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment method").SiebGetRoProperty("activeitem")
				If sPaymentMethod1 <> DictionaryTest_G.Item("PaymentMethod") Then
					AddVerificationInfoToResult  "Pass" , "Payment method is changed successfully from " & DictionaryTest_G.Item("PaymentMethod") & " to " & sPaymentMethod1
					CaptureSnapshot()
					Exit Function
				Else
					AddVerificationInfoToResult  "Fail" , "Payment method is not changed after performing Tokenization"
					iModuleFail = 1
					Exit Function
				End If
		Else
			AddVerificationInfoToResult  "Fail" , "Billing Profile page is not dislayed."
			iModuleFail = 1
			Exit Function
		End If

	End If

	'Flow
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3" ' clicking on Profile tab
	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").Exist(3) Then
		sPaymentMethod = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebPicklist("Payment method").SiebGetRoProperty("activeitem")
		AddVerificationInfoToResult  "Info" , "Payment method : " & sPaymentMethod & " is displayed"
	Else
		AddVerificationInfoToResult  "Fail" , "Billing Profile page is not dislayed."
		iModuleFail = 1
		Exit Function
	End If

	If DictionaryTest_G.Exists("PaymentMethod") Then
		DictionaryTest_G.Item("PaymentMethod")=sPaymentMethod
	else
		DictionaryTest_G.add "PaymentMethod",sPaymentMethod
	End If
	
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebMenu("Menu").Select "AboutRecord"
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : CatalogDirectDebit_fn
' 	Description : This function is used to capture comments in Account Summary page under Customer Comms list
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CatalogDirectDebit_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CatalogDirectDebit$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sMessage = adoData( "Message")  & ""

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "SIS OM Customer Account Portal View","L3" ' clicking on Account SUmmary page
	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "SIS OM Customer Account Portal View","L3"
	Browser("index:=0").Page("index:=0").Sync

	'Flow

		strComment = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List").SiebList("List").SiebGetCellText("Comments",0)
	If Instr(strComment,sMessage) > 0 Then
			AddVerificationInfoToResult "Pass","Comment is updated successfullly in Customer Comms tab and comment is :" & strComment
		Else
			AddVerificationInfoToResult "Fail","Comment is not updated successfully and comment is " & strComment
			iModuleFail = 1
		End If


		CaptureSnapshot()


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : ValidateOutboundEmailPressF9_fn
' 	Description : This function is used to Press F9 at Account Summary Page and validate outbound email.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ValidateOutboundEmailPressF9_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [EmailPressF9$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sToEmail = adoData( "ToEmail")  & ""
	sEmailBody = adoData( "EmailBody")  & ""
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

		Set oShell = CreateObject("WScript.Shell") ' Crate shell object, pressing F9 button
			oShell.SendKeys ("{F9}")

		If  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Recipient").SiebList("List").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Recipient").SiebList("List").ActivateRow 1
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Recipient").SiebButton("OK").SiebClick False
		Else
			AddVerificationInfoToResult  "Fail" ,"F9 button is not clicked successfully."
			iModuleFail = 1
			Exit Function
		End If

		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Send Email").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Send Email").SiebPicklist("SiebPicklist").SiebSelect "VF_Email"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Send Email").SiebText("SiebText").SiebSetText sToEmail
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Send Email").SiebText("SiebText_2").SiebSetText "Email"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Send Email").SiebPicklist("SiebPicklist_2").SiebSelect sEmailBody
			CaptureSnapshot()
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Send Email").SiebButton("Send").SiebClick False
			AddVerificationInfoToResult  "Info" ,"Details are updated successfully."
		Else
			AddVerificationInfoToResult  "Fail" ,"Send Email page is not displayed successfully after Pressing on F9 button."
			iModuleFail = 1
		End If
		
        
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : TopUpVoucher_fn
' 	Description : This function is used to click on Real Time Balance, check balance and place TopUp Request
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function TopUpVoucher_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [TopUpVoucher$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sSelectItem = adoData( "SelectItem")  & ""
	sVoucherNumber = DictionaryTest_G.Item("Voucher")
	sTopUpMethod = adoData( "TopUpMethod")  & ""
	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebScreenViews("ScreenViews").Goto "VF Balance Prepaid Billing Profile View","L3"

	rowNum = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Product / Services").SiebList("List").SiebListRowCount

	If rowNum = 0 Then
		AddVerificationInfoToResult  "Fail" , "Real Time Balance tab is not clicked successfully."
		iModuleFail = 1
		Exit Function
	End If

	res = False

	For loopVar = 0 to rowNum-1
		strVal = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Product / Services").SiebList("List").SiebGetCellText("Name",loopVar)

		If Instr(sSelectItem,strVal) > 0 Then
			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Product / Services").SiebList("List").ActivateRow loopVar
			res = True
			Exit For
		End If
	Next

	If res = False Then
		AddVerificationInfoToResult  "Fail" , "Item : " & sSelectItem & " is not present in WebList."
		iModuleFail = 1
		Exit Function
	End If

	SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebButton("Full Balance").SiebClick False

	sRemainingQuantity = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebList("List").SiebGetCellText("Remaining qty",0)

	AddVerificationInfoToResult  "Info" , "Initial Balance before Top Up is " & sRemainingQuantity

	CaptureSnapshot()

    SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Product / Services").SiebButton("Top-Up Request").SiebClick False

	If SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("Top-Up Request").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("Top-Up Request").SiebButton("New").SiebClick False   'clicking on New Button
	Else
		AddVerificationInfoToResult  "Fail" , "TopUp Request button is not clicked successfully."
		iModuleFail = 1
		Exit Function
	End If

	If SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("TopUp Management").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("TopUp Management").SiebPicklist("Topup method").SiebSelect sTopUpMethod
		SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("TopUp Management").SiebText("Voucher no.").SiebSetText sVoucherNumber
		SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("TopUp Management").SiebButton("Validate Voucher").SiebClick False
		CaptureSnapshot()
		If SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("TopUp Management").SiebButton("Submit").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("TopUp Management").SiebButton("Submit").SiebClick sPopUp
		Else
			AddVerificationInfoToResult  "Fail" , "Voucher number is not updated successfullly.	"
			iModuleFail = 1
			Exit Function	
		End If
	Else
		AddVerificationInfoToResult  "Fail" , "New Button is not clicked successfully on TopUp Request Page."
		iModuleFail = 1
		Exit Function
	End If
	
	If SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("Top-Up Request").SiebButton("Balance Details").Exist(2) Then
		 SiebApplication("Siebel Call Center").SiebScreen("Products & Services").SiebView("Top-Up Request").SiebApplet("Top-Up Request").SiebButton("Balance Details").SiebClick False
	Else
		AddVerificationInfoToResult  "Fail" , "Top up request is not placed successfully."
		iModuleFail = 1
		Exit Function		
	End If

	SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebButton("Full Balance").SiebClick False

	sAfterTopUpQuantity = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebList("List").SiebGetCellText("Remaining qty",0)

	AddVerificationInfoToResult  "Info" , "Balance after Top Up is " & sAfterTopUpQuantity

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : GotoSubAccounts_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function GotoSubAccounts_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Sub Account View","L3"
	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : GotoAccountSummary_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function GotoAccountSummary_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "SIS OM Customer Account Portal View","L3"
	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : ClickNewBillingAccount_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ClickNewBillingAccount_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Billing Accounts").SiebButton("New").SiebClick False

	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : ClickNewServiceAccount_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ClickNewServiceAccount_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Service Accounts").SiebButton("New").Click

	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : GotoAccountsHierarchy_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function GotoAccountsHierarchy_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Hierarchy Analysis List View","L3"
	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : BulkModify_fn
' 	Description :  This function is used to click on Bulk modify and perform desired action
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BulkModify_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [BulkModify$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Bulk Order Modify View","L3" ' click on Bulk modify tab on accounts page
	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Bulk Modify")


	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	
		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
		End If
		If sLocateCol <> "" Then
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" , "Promotion not found in the list."
				Exit Function
			End If
		End If
		If sUIName <> "" Then
			UpdateSiebList objApplet,sUIName,sValue
			If instr (sUIName, "CaptureTextValue1") > 0 AND instr (sUIName, "Order Number") > 0 Then
				If DictionaryTest_G.Exists("OrderNo") Then
					DictionaryTest_G.Item("OrderNo")=sValue
				else
					DictionaryTest_G.add "OrderNo",sValue
				End If
			End If
		End If
		adoData.MoveNext
	Loop


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : BulkModify_TargetPromotionProcess_ImportDialog_fn
' 	Description :  This function is used to click on click on pop up under Target promotion  Processcolumn and select the values from import dialog window.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BulkModify_TargetPromotionProcess_ImportDialog_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [BulkModify$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

'	sSelect = adoData( "Select")  & ""
	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""

	index = 0
	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Import Dialog").exist(10) Then
				Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Import Dialog")
				res=LocateColumns (objApplet,sLocateCol,sLocateColValue,index)
				If Cstr(res)<>"True" Then
					 iModuleFail = 1
					AddVerificationInfoToResult  "Error" , "Value : - " & sLocateColValue & " not found in the list."
					Exit Function
				End If	
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Import Dialog").SiebButton("OK").SiebClick False
			
				AddVerificationInfoToResult  "Pass" , "Target Promotion Process is done successfully."
		Else
				AddVerificationInfoToResult  "Fail" , "Target Promotion Process not done successfully."
					iModuleFail = 1

	End If
		CaptureSnapshot()
	
		If Err.Number <> 0 Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , Err.Description
		End If
	
	
	
End Function


'#################################################################################################
' 	Function Name : BulkModify_TargetPromotion_AllPromotions_fn
' 	Description :  This function is used to click on click on pop up under Target promotion  column and select the values from all promotions dialog window.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BulkModify_TargetPromotion_AllPromotions_fn()


Dim adoData	  
    strSQL = "SELECT * FROM [MenuBulkModify$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sPromotionName = DictionaryTest_G.Item("ProductName")
	sMenuItem = adoData( "MenuItem")  & ""

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("All Promotions").Exist(100) Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("All Promotions").SiebPicklist("SiebPicklist").SiebSelect "Promotion Name"
				If  instr(sMenuItem,"Tariff")=0 Then
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("All Promotions").SiebText("SiebText").SiebSetText sPromotionName
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("All Promotions").SiebButton("Go").Click
					AddVerificationInfoToResult  "Info" , "Target Promotion is selected successfullly from popup window."
				else 
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("All Promotions").SiebButton("OK").Click
							AddVerificationInfoToResult  "Info" , "Target Promotion is selected successfullly from popup window."
				 End If	
				AddVerificationInfoToResult  "Info" , "Target Promotion is selected successfullly from popup window."
	Else
		AddVerificationInfoToResult  "Fail" , "Target Promotion pop up is not clicked successfully."
		iModuleFail = 1
	End If
	

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function



'#################################################################################################
' 	Function Name : BulkPosttoPre_AccountSelection_fn
' 	Description :  This function is used to click on click on pop up under Target promotion  column and select the values from all promotions dialog window.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BulkPosttoPre_AccountSelection_fn()


	Dim adoData	  
	   strSQL = "SELECT * FROM [BulkModify$] WHERE  [RowNo]=" & iRowNo
		Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
		'sOR
		call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	 sAccountNumber = DictionaryTest_G.Item("PrePostAccountNo")

			Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Bulk Modify")

			
				Do while Not adoData.Eof
					sLocateCol = adoData( "LocateCol")  & ""
					sLocateColValue = adoData( "LocateColValue")  & ""
					sUIName = adoData( "UIName")  & ""
					sValue = adoData( "Value")  & ""	
					iIndex = adoData( "Index")  & ""
					If iIndex="" Then
						iIndex=0
					End If
					If sLocateCol <> "" Then
						res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
						If  Cstr(res)<>"True" Then
							 iModuleFail = 1
							AddVerificationInfoToResult  "Error" , "Promotion not found in the list."
							Exit Function
						End If
					End If
					If sUIName <> "" Then
						UpdateSiebList objApplet,sUIName,sValue
						End If

								Browser("index:=0").Page("index:=0").Sync
								If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Account").Exist(100)  Then
									SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Account").SiebText("SiebText").SetText sAccountNumber
									SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
					
					'				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").SetText 				
					'				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Account").SiebButton("OK").Click
									Browser("index:=0").Page("index:=0").Sync
			
									rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Account").SiebList("List").SiebListRowCount
											If rowCnt <> 0 Then
												SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
												Else
													iModuleFail = 1
													AddVerificationInfoToResult "Fail","No rows displayed after clicking on Go button."
													Exit Function
											End If
									AddVerificationInfoToResult  "Info" , "Target Account is selected successfullly from popup window."
							Else
									AddVerificationInfoToResult  "Fail" , "Target Account pop up is not clicked successfully."
									iModuleFail = 1
							End If
				
					adoData.MoveNext
				Loop
			
			
			

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : MenuBulkModify_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function MenuBulkModify_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [MenuBulkModify$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""
	sIndex = adoData( "Index")  & ""
	If sIndex="" Then
		sIndex=0
	End If
	sMenuItem = adoData( "MenuItem")  & ""
	sModify = adoData( "Modify")  & ""


	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Bulk Modify")

	res=LocateColumns (objApplet,sLocateCol,sLocateColValue,sIndex)
	If Cstr(res)<>"True" Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Value : - " & sLocateColValue & " not found in the list."
		Exit Function
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Bulk Modify").SiebMenu("Menu").SiebSelectMenu sMenuItem

	wait 5

	If sModify = "Yes" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Bulk Modify").SiebButton("Modify").SiebClick False
			If Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").Exist(5) Then ' checking for PopUp
				Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").WebButton("Confirm").Click
			End If
	End If

    
	CaptureSnapshot()

	wait 150

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Bulk Modify").SiebMenu("Menu").SiebSelectMenu "Run Query"

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : BulkViewProcessStatusValidation_fn
' 	Description :  This function is used to click on Bulk modify and perform desired action
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function BulkViewProcessStatusValidation_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [BulkModify$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Browser("index:=0").Page("index:=0").Sync

		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").GetRoProperty("ActiveScreen") <>  "VF Bulk Order Modify View" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Bulk Order Modify View","L3" 
		End If

	' click on Bulk modify tab on accounts page
	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Bulk Modify")

				sStatus =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Bulk Modify").SiebList("List").SiebText("Process Status").SiebGetRoProperty("text")
				i = 0
		
						Do while sStatus <> "Completed" and i <=50			
							wait (5)			
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Bulk Modify").SiebMenu("Menu").Select  "ExecuteQuery"		
								Do while Not adoData.Eof
										sLocateCol = adoData( "LocateCol")  & ""
										sLocateColValue = adoData( "LocateColValue")  & ""
										sUIName = adoData( "UIName")  & ""
										sValue = adoData( "Value")  & ""	
										iIndex = adoData( "Index")  & ""
										If iIndex="" Then
											iIndex=0
										End If
										If sLocateCol <> "" Then
											res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
											CaptureSnapshot()
											If  Cstr(res)<>"True" Then
												 iModuleFail = 1
												AddVerificationInfoToResult  "Error" , "Row not found in the list."
												CaptureSnapshot()
												Exit Function
											End If
										End If
										If sUIName <> "" Then
											UpdateSiebList objApplet,sUIName,sValue
										
										End If
										adoData.MoveNext
									Loop	

							sStatus =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Bulk Modify").SiebList("List").SiebText("Process Status").SiebGetRoProperty("text")
					
							If sStatus = "Completed" Then
										AddVerificationInfoToResult  "Error" , "Process Status is Complete"
								Exit Do
							End If			
							i = i+1			
						Loop
		
					If   sStatus = "Completed"Then
							AddVerificationInfoToResult  "Error" , "Process Status is Complete"
							else
							  iModuleFail = 1
								AddVerificationInfoToResult  "Error" , "Process Status is : "&sStatus
					End If


			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If
		
		
		End Function


'#################################################################################################
' 	Function Name : BulkModifyAddProducts_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BulkModifyAddProducts_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [BulkModifyAddProducts$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

		sModify = adoData( "Modify")  & ""
		sProductName = adoData( "ProductName")  & ""
		sTransferOfOwnership = adoData("TransferOfOwnership")  & ""
		sProcessStatus = adoData("ProcessStatus")  & ""
		sPopUp = adoData( "Popup")  & ""
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

			If  sModify = "N" Then
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Available").Exist(10) Then
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Available").SiebPicklist("SiebPicklist").SiebSelect "Product Name"
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Available").SiebText("SiebText").SiebSetText sProductName
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Available").SiebButton("Go").SiebClick False
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Selected").SiebButton("Add >").SiebClick False
					CaptureSnapshot()
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary_2").SiebApplet("Selected").SiebButton("OK").SiebClick False
				Else
					AddVerificationInfoToResult  "Fail" ,"Open Pop Up is not clicked successfully."
					iModuleFail = 1
				End If
		End If

		If sModify = "Y" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Bulk Modify").SiebButton("Modify").SiebClick False


				If Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").Exist(3) Then
					Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").WebButton("Confirm").Click
				End If
			Browser("index:=0").Page("index:=0").Sync
			End If

		If  sTransferOfOwnership = "Y" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Bulk Modify").SiebButton("Transfer Of OwnerShip").SiebClick sPopUp
				Browser("index:=0").Page("index:=0").Sync
				CaptureSnapshot()
		End If



	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : AnonymousAddress_fn
' 	Description : This function is used to used to check 
'   Created By :  Tarun
'	Creation Date :        
'#####################################################################correct anonymous address.#############################
Public Function AnonymousAddress_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AnonymousAddress$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sEditable = adoData( "Editable")  & ""
	sExpectedAddressVal = adoData( "AddressValue")  & ""

	'Flow
	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Address Line 1").highlight

	If sEditable = "Address Line 1" Then

		sAddressVal = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Address Line 1").GetRoProperty("text"))

		If sAddressVal = sExpectedAddressVal Then
			AddVerificationInfoToResult "Info","Address Value is : " & sAddressVal & " for anonymous customer for Address Line 1."
		Else
			AddVerificationInfoToResult "Fail","Address Value is : " & sAddressVal & " for anonymous customer but it should be : " & sExpectedAddressVal
			iModuleFail = 1
		End If
		CaptureSnapshot()
	End If


	If sEditable = "Profiles" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"

		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile")

			Do while Not adoData.Eof
				sLocateCol = adoData( "LocateCol")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""	
				iIndex = adoData( "Index")  & ""
				If iIndex="" Then
					iIndex=0
				End If
				If sLocateCol <> "" Then
					res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
					If  Cstr(res)<>"True" Then
						 iModuleFail = 1
						AddVerificationInfoToResult  "Error" , "Promotion not found in the list."
						Exit Function
					End If
				End If
				If sUIName <> "" Then
					UpdateSiebList objApplet,sUIName,sValue
					If Trim(sValue) = sExpectedAddressVal Then
						AddVerificationInfoToResult "Info","Address Value is : " & sValue & " for anonymous customer under Profiles tab."
					Else
						AddVerificationInfoToResult "Fail","Address Value is : " & sValue & " for anonymous customer but it should be : " & sExpectedAddressVal
						iModuleFail = 1
					End If

				End If
				adoData.MoveNext
			Loop
			CaptureSnapshot()
	End If

		If sEditable = "Credit Vetting" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Credit Vetting View","L3"

		If Browser("Siebel Call Center").Dialog("Message from webpage").Static("Please enter a minimum").Exist(5) Then
			Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
		End If

		If Browser("Siebel Call Center").Dialog("Message from webpage").Static("Please enter additional").Exist(3) Then
			Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
		End If

		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities").SiebApplet("Address History")

			Do while Not adoData.Eof
				sLocateCol = adoData( "LocateCol")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""	
				iIndex = adoData( "Index")  & ""
				If iIndex="" Then
					iIndex=0
				End If
				If sLocateCol <> "" Then
					res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
					If  Cstr(res)<>"True" Then
						 iModuleFail = 1
						AddVerificationInfoToResult  "Error" , "Promotion not found in the list."
						Exit Function
					End If
				End If
				If sUIName <> "" Then
					UpdateSiebList objApplet,sUIName,sValue
					If Trim(sValue) = sExpectedAddressVal Then
						AddVerificationInfoToResult "Info","Address Value is : " & sValue & " for anonymous customer under Credit Vetting tab."
					Else
						AddVerificationInfoToResult "Fail","Address Value is : " & sValue & " for anonymous customer but it should be : " & sExpectedAddressVal
						iModuleFail = 1
					End If

				End If
				adoData.MoveNext
			Loop
			CaptureSnapshot()
	End If

	If sEditable = "Addresses" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Business Address View","L3"
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses")


			Do while Not adoData.Eof
				sLocateCol = adoData( "LocateCol")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""	
				iIndex = adoData( "Index")  & ""
				If iIndex="" Then
					iIndex=0
				End If
				If sLocateCol <> "" Then
					res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
					If  Cstr(res)<>"True" Then
						 iModuleFail = 1
						AddVerificationInfoToResult  "Error" , "Promotion not found in the list."
						Exit Function
					End If
				End If
				If sUIName <> "" Then
					UpdateSiebList objApplet,sUIName,sValue
					If Trim(sValue) = sExpectedAddressVal Then
						AddVerificationInfoToResult "Info","Address Value is : " & sValue & " for anonymous customer under Addresses tab."
					Else
						AddVerificationInfoToResult "Fail","Address Value is : " & sValue & " for anonymous customer but it should be : " & sExpectedAddressVal
						iModuleFail = 1
					End If

				End If
				adoData.MoveNext
			Loop
			CaptureSnapshot()
	End If



	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : GotoAgreements_fn
' 	Description : This function is used to click on PAgreements tab on Account page 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function GotoAgreements_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Agreement List View","L3"
	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : CreateNewServiceRequest_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By :  Ankit/Tarun
'	Creation Date :        
'##################################################################################################
Public Function CreateNewServiceRequest_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""
		sMenuItem =  adoData( "SelectMenu")  & ""
	sClickQuery =  adoData( "ClickQuery")  & ""
	sSave =  adoData( "Save")  & ""

		sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Service Account List View","L3"

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").getROProperty("activeview") <> "Service Account List View" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Service Account List View","L3"
	End If

	Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If sSave = "Yes"  Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Account Details").SiebButton("Save").SiebClick sPopUp
		Exit Function
	End If

	If Ucase(sClickNew) = "NO" Then
	Else
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Service Requests").SiebButton("New").SiebClick False
	End If
	If Ucase(sClickQuery) = "YES" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Service Requests").SiebButton("Query").SiebClick False
	End If

	Set obj = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Service Requests")

	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Service Requests").SiebList("List").ActivateRow 0
	CaptureSnapshot()

	If sMenuItem <> "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Service Requests").SiebMenu("Menu").SiebSelectMenu sMenuItem
	End If

	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	
		sValue = Replace (sValue,"SRID",DictionaryTempTest_G("SRID"))
		sClickGo =  adoData( "ClickGo")  & ""

		UpdateSiebList obj,sUIName,sValue
		If instr (sUIName, "CaptureTextValue") > 0 Then
			If  instr (sUIName, "SRRowId") > 0 Then
				DictionaryTempTest_G("SRRowId")=sValue
			elseif  instr (sUIName, "ID") > 0 Then
				DictionaryTempTest_G("SRID")=sValue
			End If			
		End If
		CaptureSnapshot()
		If Ucase(sClickGo) = "YES" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Service Requests").SiebButton("Go").SiebClick False
		End If
		adoData.MoveNext
	Loop
		CaptureSnapshot()
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : ServiceRequestScreen_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ServiceRequestScreen_fn()

	'Get Data
	Dim adoData	  
	strSQL = "SELECT * FROM [CreateNewServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Service Request Screen"

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Service Request Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Service Request Screen"
	End If

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Service Request Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Service Request Screen"
	End If

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebScreenViews("ScreenViews").Goto "Personal Service Request List View","L2"

	Browser("index:=0").Page("index:=0").Sync


	If Ucase(sClickNew) = "NO" Then
	Else
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebButton("New").SiebClick False
	End If
	CaptureSnapshot()
	
	SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").ActivateRow 0

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests")

	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
				If sValue = "NOW" Then
					TimeRequested = FormatDateTime (Time, vbShortTime)& Right (Time, 6)
					MN =Left( MonthName(Month(Date)),3)
					DateRequested = Date					
					DateRequested = day(DateRequested) & "/" & MN & "/" &  year(DateRequested) &" "& TimeRequested
					sValue = DateRequested								
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
' 	Function Name : ServiceRequestBelowBillingProfile_fn
' 	Description :  This function is used to click on new button in Service accounts applet on account s page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ServiceRequestBelowBillingProfile_fn()


	Dim adoData	  
	strSQL = "SELECT * FROM [CreateNewServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sClickNew = adoData( "ClickNew")  & ""
	sDrillDown = adoData( "DrillDown")  & ""

	'Flow

	Browser("index:=0").Page("index:=0").Sync
    
	If Ucase(sClickNew) = "NO" Then
	Else
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").SiebButton("New").SiebClick False
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").SiebList("List").ActivateRow 0

    Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests")

	 If Ucase(sNonEditable) = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").SiebList("List").SiebClick False
		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").SiebList("List").SiebPicklist("Request Source").isenabled = False Then
				 CaptureSnapshot()
                 AddVerificationInfoToResult "Pass","Request Source is Disabled as Expected"
		Else
				iModuleFail = 1
				AddVerificationInfoToResult "Fail","Request Source is Enabled"	
        End If

	End If


	Do while Not adoData.Eof

		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""

		If sUIName <> "" Then		
			UpdateSiebList objApplet,sUIName,sValue
		End If

		adoData.MoveNext
	Loop

	If Instr(sUIName,"OpenPopUp|Installed ID") > 0 Then
		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebButton("OK_2").SiebClick False
			AddVerificationInfoToResult  "Info" , "Installed Id is selected successfully."
		End If
	End If

'	If sDrillDown = "Yes" Then
'		On error resume next
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").SiebMenu("Menu").Select "WriteRecord"
'
'		
'		err.clear
'
'		If Browser("Siebel Call Center").Dialog("Message from webpage").Static("Please ensure you have").Exist(5) Then
'			Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
'		End If
'	End If

	If Instr(sUIName,"OpenPopUp|Owner") > 0 Then
		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebPicklist("SiebPicklist").SiebSelect "User ID"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebText("SiebText").SiebSetText "test_agent_nba_04"
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebButton("Go").SiebClick False
			AddVerificationInfoToResult  "Info" , "User login is changed from test_retail_1 to test_agent_nba_04"
		End If
	End If

	CaptureSnapshot()
'	err.clear
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function




''#################################################################################################
' 	Function Name : VerifyOnlineAnonymousAccount_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function VerifyOnlineAnonymousAccount_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [OnlineAnonymous$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	sFirst_Name = adoData( "First_Name")  & ""
	sLast_Name = adoData( "Last_Name")  & ""
	sVerifyChkboxDisabled = adoData( "VerifyChkboxDisabled")  & ""
	sVerifyChkboxEnabled = adoData( "VerifyChkboxEnabled")  & ""
	sAccountName= sFirst_Name & " " & sLast_Name
	Browser("index:=0").Page("index:=0").Sync
	actAccountName=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebGetCellText ("Account name", 0)
	actAccountNo=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebGetCellText ("Account no.", 0)

				If DictionaryTest_G.Exists("AccountNo") Then
					DictionaryTest_G.Item("AccountNo")=actAccountNo
				else
					DictionaryTest_G.add "AccountNo",actAccountNo
				End If
			
				AddVerificationInfoToResult  "Info" , "AccountNo is " & actAccountNo

		If sVerifyChkboxDisabled = "Y" Then
					CreateAccChkBox = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Create online account").GetROProperty("isenabled")
					CaptureSnapshot()
				
					If CreateAccChkBox = "False" Then
						AddVerificationInfoToResult "Pass","Create OnlineAccount Checkbox is Disabled"
					else
						  iModuleFail = 1
						AddVerificationInfoToResult "Fail","Create OnlineAccount Checkbox is not Disabled"
					End If
				
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
					Browser("index:=0").Page("index:=0").Sync
	
					AccSummChkBox = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebCheckbox("Online account").GetROProperty("isenabled")
					CaptureSnapshot()
					If AccSummChkBox = "False" Then
						AddVerificationInfoToResult "Pass","OnlineAccount Checkbox on Account Summary is Disabled"
					else
						  iModuleFail = 1
						AddVerificationInfoToResult "Fail","OnlineAccount Checkbox on Account Summary  is not Disabled"
					End If
	
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
					Browser("index:=0").Page("index:=0").Sync
					CaptureSnapshot()
					If sVerifyChkboxEnabled = "Y" Then
						AccContChkBox = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").SiebCheckbox("Online ID").GetROProperty("isenabled")
						CaptureSnapshot()
								If AccContChkBox = "True" Then
									AddVerificationInfoToResult "Pass","Online ID Checkbox is Disabled"
								else
									  iModuleFail = 1
									AddVerificationInfoToResult "Fail","Online ID Checkbox is Not  Disabled"
								End If
					Else
							AccContChkBox = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").SiebCheckbox("Online ID").GetROProperty("isenabled")
							CaptureSnapshot()
							If AccContChkBox = "False" Then
								AddVerificationInfoToResult "Pass","Online ID Checkbox is Disabled"
							else
								  iModuleFail = 1
								AddVerificationInfoToResult "Fail","Online ID Checkbox is Not  Disabled"
							End If
				End If
				
	End If

End Function

'#################################################################################################
' 	Function Name : InstalledIDSelection_fn
' 	Description : This function is used to used to click on Profile tab on Account page,capture Payment method  and select About Record after selecting billing profile.
'   Created By :  
'	Creation Date :        
'##################################################################################################
Public Function InstalledIDSelection_fn()


		'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Pick Asset").SiebButton("OK").SiebClick False
'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Pick Asset").SiebButton("OK").SiebClick False

		 SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Pick Asset").SiebButton("OK").SiebClick False
		
		 CaptureSnapshot()

End Function
'#################################################################################################
' 	Function Name : AgeIDVerification_fn()
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function AgeIDVerification_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AgeIDVerification$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sVerificationType = adoData("VerificationType") & ""
	sApplicationType = adoData("ApplicationType") & ""
	sOverrideVerification = adoData("OverrideVerification") & ""
	sOutcome = adoData("Outcome") & ""
    sAgeVerified = adoData("AgeVerified") & ""
    sAddressType = adoData("AddressType") & ""
    sIDType = adoData("IDType") & ""
	sContactTobeAgeVerified = adoData("ContactTobeAgeVerified") & ""

	sPopUp = adoData( "Popup")  & ""

	If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If


	''Flow
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Age Verification View","L3"
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebPicklist("Verification type").SiebSelect sVerificationType
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebPicklist("Application Type").SiebSelect sApplicationType


		If  sAgeVerified = "Y" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebText("Contact to be age verified").OpenPopup
			Browser("index:=0").Page("index:=0").Sync
				If  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Contact").exist(5)Then
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Pick Contact").SiebButton("OK").Click
					AddVerificationInfoToResult "Pass" , "Pick Contact Applet shown successfully and Account selected successfully "
				Else
					AddVerificationInfoToResult "Fail" , "Pick Contact Applet not shown"
					iModuleFail = 1
				End If
			
			
		End If
		

		If sIDType<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebPicklist("Proof of ID type").SiebSelect sIDType
		End If

		If  sAddressType<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebPicklist("Proof of address type").SiebSelect sAddressType
		End If	

		If sOverrideVerification<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebCheckbox("Override verification").SiebSetCheckbox sOverrideVerification
		End If

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebButton("Submit").SiebClick sPopUp
		Browser("index:=0").Page("index:=0").Sync

		If  sOutcome <>"" Then
			sOutcomeActual = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Account Details").SiebText("Outcome").GetROProperty("text")
			CaptureSnapshot()
				If Trim(sOutcomeActual) = Trim(sOutcome) Then
					AddVerificationInfoToResult "Info" , "Outcome is " &sOutcomeActual
				Else
					AddVerificationInfoToResult "Fail" , "Outcome is" &sOutcomeActual
				iModuleFail = 1
			End If
		End If
		
		
		CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	

End Function

''#################################################################################################
'               Function Name : CreditVettingTabVerify_fn()
'               Description : 
'   Created By :  Pushkar Vasekar
'               Creation Date :        
'##################################################################################################
Public Function CreditVettingTabVerify_fn()

'Get Data
                Dim adoData        
    strSQL = "SELECT * FROM [CreditVetting$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sBankAccountName = DictionaryTest_G.Item("Name")
	sBankAccountNo = DictionaryTest_G.Item("Accno")
	sKeepBankDetails = adoData("KeepBankDetails") & ""
	sSortCode = DictionaryTest_G.Item("Sortcode")
	sTransactCustomerType = adoData("TransactCustomerType") & ""
	sResidentialStatus = adoData("ResidentialStatus") & ""
	sEmploymentStatus = adoData("EmploymentStatus") & ""
	sBillPaymentMethod = adoData("BillPaymentMethod") & ""
	sUseExistingDetails = adoData("UseExistingDetails") & ""
	sNewButton = adoData("NewButton") & ""
	sPostCode = adoData("PostCode") & ""
	sCreateOrder = adoData("CreateOrder") & ""


	sPopUp = adoData( "Popup")  & ""
	sPopUp1 = adoData( "Popup1")  & ""
	sPopUp2 = adoData( "Popup2")  & ""

	If Ucase(sPopUp)="FALSE" or sPopUp=""Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If

	If Ucase(sPopUp1)="FALSE" or sPopUp1=""Then
		sPopUp1="FALSE"
		sPopUp1=Cbool(sPopUp1)
	End If
	If Ucase(sPopUp2)="FALSE" or sPopUp2=""Then
		sPopUp2="FALSE"
		sPopUp2=Cbool(sPopUp2)
	End If
	If Ucase(sPopUp3)="FALSE" or sPopUp3=""Then
		sPopUp3="FALSE"
		sPopUp3=Cbool(sPopUp3)
	End If

''			Flow

			If sNewButton = "Y" Then

                	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Address History").SiebButton("New").isenabled = True then
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Address History").SiebButton("New").click
							AddVerificationInfoToResult "Info" , "New Button is clicked successfully "
					Else
							AddVerificationInfoToResult "Fail" , "New Button is not enabled "
							 iModuleFail = 1
							 Exit Function
					End If


					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SiebSetText sPostCode
					CaptureSnapshot()
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").ProcessKey "EnterKey"
					Browser("index:=0").Page("index:=0").Sync
					SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Add Address").SiebButton("OK").Click
					Browser("index:=0").Page("index:=0").Sync
					HandlePopup(sPopUp3)
					Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Address History")
							objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

					Do while Not adoData.Eof
			
					sUIName = adoData( "UIName")  & ""
					sValue = adoData( "Value")  & ""
					If sValue = "Backdate" Then
						DateRequested = Date   
						MN =Left( MonthName(Month(Date)),3)                                     
						sValue = day(DateRequested) & "/" & MN & "/" & (year(DateRequested))-3 &" "& TimeRequested
					End If
    		
					iIndex = adoData( "Index")  & ""
					If iIndex="" Then
						iIndex=0
					End If
			
					If sUIName <> "" Then		
						UpdateSiebList objApplet,sUIName,sValue
					End If
					CaptureSnapshot()
					adoData.MoveNext
			
				Loop

                	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Address History").SiebMenu("Menu").Select "Run Query"
					If  Browser("Siebel Call Center_2").Dialog("Message from webpage").Exist(5) Then
							CaptureSnapshot()
                            AddVerificationInfoToResult "Fail" , "POPUP came as the total of all address is not more than 3 yrs"
							iModuleFail = 1
							Exit Function
					Else
							AddVerificationInfoToResult "PASS" , "POPUP did not came as the total of all the address is more than 3 yrs "
    				End If	


        	Else

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Credit Vetting View","L3"
            CaptureSnapshot()
			
			
			if Browser("Siebel Call Center_2").Dialog("Message from webpage").Static("IF A CREDIT VET IS REQUIRED").Exist(5) Then
					CaptureSnapshot()
					Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Click
					AddVerificationInfoToResult "Info" , "POPUP came as expected"
			Else
					AddVerificationInfoToResult "Fail" , "POPUP did not came "
					 iModuleFail = 1
					 Exit Function
			End If

			CaptureSnapshot()
            HandlePopup(sPopUp)

			If BankAccountNo <> "" Then
						Browser("index:=0").Page("index:=0").Sync
						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities").SiebApplet("Credit Vetting").SiebText("Bank account no").SiebSetText sBankAccountNo
						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities").SiebApplet("Credit Vetting").SiebText("Bank account name").SiebSetText sBankAccountName
						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities").SiebApplet("Credit Vetting").SiebText("Sort code").SiebSetText sSortCode
								
						On error resume next
						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities").SiebApplet("Credit Vetting").SiebCheckbox("Keep my bank details for").SetOn 
						err.clear
			End If

			If  sUseExistingDetails = "Y"Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Credit Vetting").SiebText("Use existing DD billing<br/>pr").OpenPopup
				Browser("index:=0").Page("index:=0").Sync
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Billing profile").SiebButton("OK").SiebClick sPopUp

					If  Browser("Siebel Call Center_2").Dialog("Message from webpage").Static("IF A CREDIT VET IS REQUIRED").Exist(2) Then
							CaptureSnapshot()
							Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Click
					End If
'				HandlePopup(sPopUp)
            End If



'			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities").SiebApplet("Credit Vetting").SiebPicklist("Transact customer type").SiebSelect "Consumer"
			CaptureSnapshot()
			HandlePopup(sPopUp1)
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Credit Vetting_2").SiebPicklist("Transact customer type").SiebSelect sTransactCustomerType
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Credit Vetting_2").SiebPicklist("Residential status").SiebSelect sResidentialStatus
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Credit Vetting_2").SiebPicklist("Employment status").SiebSelect sEmploymentStatus 
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Credit Vetting_2").SiebPicklist("Bill payment method").SiebSelect sBillPaymentMethod
			CaptureSnapshot()

				Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Address History")
				objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

				Do while Not adoData.Eof
			
					sUIName = adoData( "UIName")  & ""
					sValue = adoData( "Value")  & ""
					If sValue = "Backdate" Then
						DateRequested = Date   
						MN =Left( MonthName(Month(Date)),3)                                     
						sValue = day(DateRequested) & "/" & MN & "/" & (year(DateRequested))-3 &" "& TimeRequested
					End If
    		
					iIndex = adoData( "Index")  & ""
					If iIndex="" Then
						iIndex=0
					End If
			
					If sUIName <> "" Then		
						UpdateSiebList objApplet,sUIName,sValue
					End If
					CaptureSnapshot()
					adoData.MoveNext
			
				Loop
				If sCreateOrder <> "" Then
								SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Credit Vetting").SiebButton("Create Order").SiebClick False
								'HandlePopup(sPopUp2)
								CaptureSnapshot()
				End If
	End If


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function

'#################################################################################################
' 	Function Name : AddDeleteMediaType_fn
' 	Description : This function adds a media type and then deletes the same to verify that at least one media type is present
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function AddDeleteMediaType_fn()
   	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AddDeleteMedia$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	call SetObjRepository ("Account",sProductDir & "Resources\")

	sMediaType = adoData( "MediaType")  & ""
	sPopUp = adoData( "Popup")  & ""
	sAction = adoData( "Action")  & ""

''
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
	Browser("index:=0").Page("index:=0").Sync
    
		 	If Ucase(sAction)="DELETE" Then ''' To check PopUp Verification for Delete 
		
		
			If Ucase(sPopUp)="FALSE" or sPopUp=""Then
				sPopUp="FALSE"
				sPopUp=Cbool(sPopUp)
			End If
			
			'sOR
			
		
			   
		
					RowC = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type").SiebList("List").SiebListRowCount
					For i = 0 to RowC-1
		
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type").SiebList("List").ActivateRow i
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type").SiebButton("New").Click
							Browser("index:=0").Page("index:=0").Sync				
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type").SiebList("List").SiebPicklist("Bill format").SiebSelect sMediaType
		'					SiebApplication("Siebel Call Center").SiebToolbar("HIQuery").Click
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type").SiebButton("Delete").SiebClick sPopUp
		
						Next
							If Not( Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("Delete").Exist) Then
								  iModuleFail = 1
								AddVerificationInfoToResult  "Error" , Err.Description
							End If
							Exit Function
		End If
	''''''	To Add new Medita Type or Edit existing one

	    Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type")

			Do while Not adoData.Eof
				sUIName = adoData( "UIName")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""
				sValue = adoData( "Value")  & ""
				sAction = adoData( "Action")  & ""

					
				If sLocateCol <> "" Then
					res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
					If  Cstr(res)<>"True" Then
'						 iModuleFail = 1
					AddVerificationInfoToResult  "Info" , sLocateCol & "-" & sLocateColValue & " not found in the list. Hence updating List"
'						Exit Function
						If Ucase(sAction)="NEW"  Then
'							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type").SiebButton("New").SiebClick
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Media Type").SiebButton("New").Click
							AddVerificationInfoToResult  "New Button" , " New Button clicked."
						End If
							If sUIName <> "" Then
						 UpdateSiebList objApplet,sUIName,sValue
						CaptureSnapshot()
							End If
					End If
				End If	 
			AddVerificationInfoToResult  "Success" , sLocateCol & "-" & sLocateColValue & " found in the list."
				
			adoData.MoveNext
			Loop ''

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : ServiceAccountVerify_fn()
' 	Description : This function is used to validate all the products on Product/Services tab in Accounts page.
'   Created By :  Payel Biswas
'	Creation Date :        
'##################################################################################################

Public Function ServiceAccountVerify_fn()

	Dim rwNum
	Dim loopVar
	Dim strProduct
	Dim sProduct
	Dim adoData
	Dim strSQL

		  
    strSQL = "SELECT * FROM [ProductServices$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")


	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Service Accounts").SiebList("List").Exist(2) Then
		AddVerificationInfoToResult "Info","Service Account list is displayed successfully in Sub Account Tab."
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Fail","Service Account  list is not displayed after clicking on Sub Account Tab on Accounts page."
		Exit Function
	End If
	


	'SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Products and Services").SiebButton("ToggleListRowCount").Click ' clicking on Toggle button to expand
	'Call ShowMoreButton_fn()

	Browser("index:=0").Page("index:=0").Sync


	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Service Accounts")
 
	Do while Not adoData.Eof
		sLocateColExpand = adoData( "LocateColExpand")  & ""
		sLocateColExpandValue = adoData( "LocateColExpandValue")  & ""
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""	
		sCollapse = adoData( "Collapse")  & ""
		sExist = adoData( "Exist")  & ""
		iIndex = 0
		If sLocateColExpand <> "" Then
			res=LocateColumns (objApplet,sLocateColExpand,sLocateColExpandValue,iIndex)
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Service Accounts").ExpandRow
		End If
		If sLocateCol <> "" Then
		
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			If sExist = "" Then
				If  res=True Then
					 AddVerificationInfoToResult  "Pass" , sLocateCol & " - " & sLocateColValue & " found in the list as expected"
				else
					 iModuleFail = 1
					AddVerificationInfoToResult  "Error" , sLocateCol & " - " & sLocateColValue & " not found in the list."				
				End If 
			else
				If sExist = "False" Then
					sExist = "False-Row Not Exist"
				End If
				If Ucase(cstr(res)) = Ucase (sExist) Then
					AddVerificationInfoToResult  "Pass" , sLocateCol & " - " & sLocateColValue & " existence is " & sExist & " as expected"
				else
					 iModuleFail = 1
					AddVerificationInfoToResult  "Pass" , sLocateCol & " - " & sLocateColValue & " existence is " & cstr(res) & " but expected is " & sExist
				End If
			End If
		End If


		If sCollapse="Y" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Service Accounts").SiebMenu("Menu").SiebSelectMenu "Run Query"
'			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectCollapseRow sCollapse,iIndex
		End If
		adoData.MoveNext
	Loop

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : CreateNewBillingProfileSubAccounts_fn
' 	Description : This function creates prepay billing profile in siebel in sub accounts
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CreateNewBillingProfileSubAccounts_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CreateNewBillingProfile$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)


	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sNewBillingProfile = adoData( "NewBillingProfile")  & ""

	'Flow

	'*****************************************************************************
	'Check billing profile screen and create prepay profile

	If SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Billing Profiles").Exist Then
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Billing Profiles").SiebButton("New").SiebClick False
	Else
		AddVerificationInfoToResult  "Error" , "Sub Accounts Billing Profile is not displayed successfully."
		 iModuleFail = 1
		 Exit Function
	End If


		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Billing Profiles")
			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

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

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : OwnedProductServices_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Payel Biswas
'	Creation Date :        
'##################################################################################################

Public Function OwnedProductServices_fn()


	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [UsedProductServices$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
'	sSelect = adoData("Select") & ""
	sAction = adoData("Action") & ""
	sEnable= adoData("Enable") & ""
    sEnableAction = adoData("EnableAction") & ""

	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""
	sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
	sLocateColValue = Replace(sLocateColValue,"InstalledId",DictionaryTest_G.Item("ACCNTMSISDN"))
'	sLocateColValue = Replace(sLocateColValue,"InstalledId","447387955408")
	sUIName = adoData( "UIName")  & ""
	sValue = adoData( "Value")  & ""

	iIndex = adoData( "Index")  & ""
	If iIndex="" Then
		iIndex=0
	End If

'SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service").SiebList("List").DrillDownColumn

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service")
'	Call ShowMoreButtonUsedProductsServices_fn()
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
		If sValue = "ID" Then
			sValue = Trim(Replace(sValue,"ID",DictionaryTest_G.Item("AgreementId")))
		End If
		UpdateSiebList objApplet,sUIName,sValue
	End If

	If (sEnable="Yes" and sValue = "Reconnection") Then
					Set menu = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service").SiebMenu("Menu")
					menuItem=menu.GetRepositoryName("Reconnection")

					On Error Resume Next
					Recovery.Enabled=False

							If menu.IsEnabled(menuItem)=True Then
							AddVerificationInfoToResult  "Pass" , "Reconnection Button is Enabled as Expected with This Login"
							Exit function

							else
							iModuleFail = 1
							AddVerificationInfoToResult "Fail","Reconnection Button is Disabled"
							Exit function
			
						End If
	End If

		If (sEnable="No"and sValue = "Reconnection") Then
					Set menu = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service").SiebMenu("Menu")
					menuItem=menu.GetRepositoryName("Reconnection")

					On Error Resume Next
					Recovery.Enabled=False

							If menu.IsEnabled(menuItem)=False Then
									AddVerificationInfoToResult  "Pass" , "Reconnection Button is Disabled as Expected with This Login"
									Exit function
							else
									iModuleFail = 1
									AddVerificationInfoToResult "Fail","Reconnection Button is Enabled"
									Exit function

							End If
		End If


	If(sAction <> "" )Then 
		If(sAction="Modify" or sAction="ModifyBar")Then 
	
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service").SiebButton("micClass:=SiebButton", "repositoryname:="&sAction).SiebClick False
	
			If Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").Exist(5) Then ' checking for PopUp
				Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").WebButton("Confirm").Click
			End If
		
		Else
	
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service").SiebMenu("Menu").SiebSelectMenu sAction
			CaptureSnapshot()

			If Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").Exist(5) Then ' checking for PopUp
				Browser("Siebel Call Center").Window("Address Change Alert ..").Page("Address Change Alert").WebButton("Confirm").Click
			End If
	
		End If

		If sValue="Reconnection" Then
					
					If   Browser("Siebel Call Center").Dialog("Siebel").Static("The MSISDN you are trying").Exist(5) Then
						 CaptureSnapshot()
						AddVerificationInfoToResult  "Pass" , "Expected pop is for 90days violation is  present."

						Else
					  iModuleFail = 1
					  AddVerificationInfoToResult  "Error" , "Expected pop is for 90days violation pop is not present."
					  Exit Function	
					End If
			End If
	
		If (sAction="Disconnect") OR (sAction="Post to Pre Migration") OR (sAction="Pre to Post Migration") Then
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Messages").Exist(5) Then
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Messages").SiebButton("Accept").SiebClick False
			End If
		End If
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function




'#################################################################################################
'               Function Name : VerifyOrderDetails_fn
'               Description : 
'   Created By :  Pushkar
'               Creation Date :        
'##################################################################################################
Public Function VerifyOrderDetails_fn()

                Dim adoData        
    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
                Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

                'sOR
                call SetObjRepository ("Account",sProductDir & "Resources\")

                sCreatedBy = adoData( "CreatedBy")  & ""
                sChannel = adoData( "Channel")  & ""
                
                
                If sCreatedBy <> "" and sChannel <> "" Then

                                sActCreatedBy = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Catalogue").SiebApplet("Orders").SiebText("Created by").GetROProperty("text")
                                sActChannel = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Catalogue").SiebApplet("Orders").SiebPicklist("Channel").GetROProperty("activeitem")

                                If Ucase(sActCreatedBy) = Ucase(sCreatedBy) and  Ucase(sActChannel) = Ucase(sChannel)  Then
                                                AddVerificationInfoToResult "Pass","Actual Value" & sCreatedBy & "matches with Expected Value" & sActCreatedBy & ""
                                                AddVerificationInfoToResult "Pass","Actual Value" & sChannel & "matches with Expected Value" & sActChannel & ""
                                else
                                iModuleFail = 1
                                AddVerificationInfoToResult "Fail","Actual Value does not match with Expected" 
                                Exit Function
                                End If
                End If
End Function


'#################################################################################################
' 	Function Name : PromotionUpgrade_fn
' 	Description : 
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################

Public Function PromotionUpgradeTV_fn()


	'Get Data
	Dim adoData
	Dim sStartingWith

    strSQL = "SELECT * FROM [PromotionUpgrade$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

'	sPromotionType = adoData("PromotionType") & ""
'	sStartingWith = adoData("StartingWith") & ""
	sStartingWith = DictionaryTest_G.Item("ProductName")

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades").Exist(1) Then
		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades")
	Else
	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades_2")
	End If
	
			'objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

			Do while Not adoData.Eof
'				sUIName = adoData( "UIName")  & ""
'				sValue = adoData( "Value")  & ""
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
'				If sUIName <> "" Then
'					 UpdateSiebList objApplet,sUIName,sValue
'					CaptureSnapshot()
'				End If
			adoData.MoveNext
			Loop ''
			CaptureSnapshot()
'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebButton("micclass:=SiebButton","repositoryname:=PopupQueryExecute").SiebClick False
	
	rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebList("micclass:=SiebList","repositoryname:=SiebList").SiebListRowCount
	If rowCnt <> 0 Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebButton("micclass:=SiebButton","repositoryname:=Execute").SiebClick False
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Fail","No rows displayed after clicking on Go button."
		Exit Function
	End If

'	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Promotion Upgrades_2").SiebButton("OK").SiebClick False
	
	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()
	'If Err.Number = 0 Then
'		Err.Clear
	'End If
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : GetAgreementDate_fn
' 	Description : Gets the start and end date for TV agreement
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################

Public Function GetAgreementDate_fn()

	'Get Data
	Dim adoData

    strSQL = "SELECT * FROM [GetAgreementDate$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Agreement List View","L3"

	'Flow
	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Agreements List View")

			Do while Not adoData.Eof
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""
				sDateDifference = adoData( "DateDifference")  & ""


					If sUIName = "DateDiff" Then
						MN =Left( MonthName(Month(Date)),3)
						DateRequested = Date					
						DateRequested = day(DateRequested) -1  & "/" & MN & "/" &  year(DateRequested)				
						sDate = DateRequested
                    DateDifference = DateDiff("m",DictionaryTest_G("StartDate"),DictionaryTest_G("EndDate"))
						If Cint(DateDifference) = Cint(sDateDifference) Then''and Left(DictionaryTest_G("StartDate"),11) = sDate 
							AddVerificationInfoToResult  "Pass" , "Agreement date difference is " & sDateDifference & " months from System date"
						Else
						 iModuleFail = 1
						AddVerificationInfoToResult  "Fail" ,"Agreement date difference is not " & sDateDifference & " months from system date"
						End If
					Exit Function
					End If

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
						 If instr (sKey1, "StartDate") > 0  Then
'						DictionaryTest_G(sKey1)=sValue
							DictionaryTest_G("StartDate")	= sValue
					End If

					 If instr (sKey1, "EndDate") > 0  Then
'						DictionaryTest_G(sKey1)=sValue
							DictionaryTest_G("EndDate")	= sValue
					End If
					CaptureSnapshot()
				End If
			adoData.MoveNext
			Loop ''
			CaptureSnapshot()
End Function


' '#################################################################################################
' ' 	Function Name : GetAgreementDate_fn
' ' 	Description : Gets the start and end date for TV agreement
' '   Created By :  Pushkar
' '	Creation Date :        
' '##################################################################################################

' Public Function GetAgreementDate_fn()

	' 'Get Data
	' Dim adoData

    ' strSQL = "SELECT * FROM [GetAgreementDate$] WHERE  [RowNo]=" & iRowNo
	' Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	' 'sOR
	' call SetObjRepository ("Account",sProductDir & "Resources\")
	' SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Agreement List View","L3"

	' 'Flow
	' Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Agreements List View")

			' Do while Not adoData.Eof
				' sUIName = adoData( "UIName")  & ""
				' sValue = adoData( "Value")  & ""
				' sLocateColValue = adoData( "LocateColValue")  & ""
				' sLocateCol = adoData( "LocateCol")  & ""
				' sDateDifference = adoData( "DateDifference")  & ""


					' If sUIName = "DateDiff" Then
						' MN =Left( MonthName(Month(Date)),3)
						' DateRequested = Date					
						' DateRequested = day(DateRequested) & "/" & MN & "/" &  year(DateRequested)				
						' sDate = DateRequested
                    ' DateDifference = DateDiff("m",DictionaryTest_G("StartDate"),DictionaryTest_G("EndDate"))
						' If Cint(DateDifference) = Cint(sDateDifference) and Left(DictionaryTest_G("StartDate"),11) = sDate Then
							' AddVerificationInfoToResult  "Pass" , "Agreement date difference is " & sDateDifference & " months from System date"
						' Else
						 ' iModuleFail = 1
						' AddVerificationInfoToResult  "Fail" ,"Agreement date difference is not " & sDateDifference & " months from system date"
						' End If
					' Exit Function
					' End If

					' iIndex = adoData( "Index")  & ""
					' If iIndex="" Then
						' iIndex=0
					' End If

				' If sLocateCol <> "" Then
					' res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
					' If  Cstr(res)<>"True" Then
						 ' iModuleFail = 1
						' AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
						' Exit Function
					' End If
				' End If	 

				' If sUIName <> "" Then
					' If instr (sUIName, "CaptureTextValue") > 0  Then
						' sKey1=sValue	
					' End If
						
					 ' UpdateSiebList objApplet,sUIName,sValue
						 ' If instr (sKey1, "StartDate") > 0  Then
' '						DictionaryTest_G(sKey1)=sValue
							' DictionaryTest_G("StartDate")	= sValue
					' End If

					 ' If instr (sKey1, "EndDate") > 0  Then
' '						DictionaryTest_G(sKey1)=sValue
							' DictionaryTest_G("EndDate")	= sValue
					' End If
					' CaptureSnapshot()
				' End If
			' adoData.MoveNext
			' Loop ''
			' CaptureSnapshot()
' End Function

'#################################################################################################
' 	Function Name : ChkUpdatedByCustComm_fn
' 	Description : This function is used to used to check updated by field in variousplaces in customer comms
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################
Public Function ChkUpdatedByCustComm_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [UpdatedByComms$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sAccountSummComms = adoData( "AccountSummComms")  & ""
	sCustomerComms = adoData( "CustomerComms")  & ""
'	sAllAcrossComms = adoData( "AllAcrossComms")  & ""
	sMyComms = adoData( "MyComms")  & ""
	sTeamComms = adoData( "TeamComms")  & ""
'	sClickNew = adoData( "ClickNew")  & ""
	sCommonComms =  adoData( "CommonComms")  & ""	
	sChckUpdatedBy = adoData( "ChckUpdatedBy")  & ""

	'Flow

		If sAccountSummComms = "Y"  Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List").SiebButton("New").SiebClick False
			Browser("index:=0").Page("index:=0").Sync
			SiebApplication("Siebel Call Center").SiebToolbar("HIQuery").Click "ExecuteQuery" ' clicking on Refresh Button
			Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List")

		ElseIf sCustomerComms = "Y" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Activities View","L3"
			Browser("index:=0").Page("index:=0").Sync
			Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities_2").SiebApplet("Customer Comms")

		ElseIf sCommonComms = "Y" Then
			sCommsID =DictionaryTest_G("CommsID")
			
			
			SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
			Browser("index:=0").Page("index:=0").Sync
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "VF All Activity List View Across My Organization","L2"
			Browser("index:=0").Page("index:=0").Sync
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Activities").SiebButton("Query").Click
			Wait 1
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Activities").SiebList("List").SiebText("ID").SiebSetText sCommsID
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Activities").SiebButton("Go").SiebClick False
			Wait 2
			Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Activities")

			
		End If

		If sMyComms = "Y" Then
			SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
			Browser("index:=0").Page("index:=0").Sync


			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Activity List View","L2"
			Browser("index:=0").Page("index:=0").Sync
			If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").Exist(1) Then
				Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms")
			Else 
				Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Activities")
			End If
			
			Browser("index:=0").Page("index:=0").Sync
		ElseIf sTeamComms = "Y" Then
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Manager's Activity List View","L2"
			If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").Exist(1) Then
				Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms")
			Else 
				Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Activities")
			End If
			Browser("index:=0").Page("index:=0").Sync
	End If

			If sChckUpdatedBy = "Y" Then
			Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Activities")
			End If

			Do while Not adoData.Eof
				objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""

				If sLocateColValue = "CommsID" Then
					sLocateColValue = DictionaryTest_G.Item("CommsID")
				End If

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
					 If instr (sUIName, "CaptureTextValue") > 0 And instr(sValue,"1-") > 0   Then
					DictionaryTest_G.Item("CommsID") = sValue
			End If

				
		End If

			adoData.MoveNext
			Loop ''

			If sChckUpdatedBy = "Y" Then
			Browser("index:=0").Page("index:=0").Sync
            sUpdatedBy = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebText("Updated by").GetRoProperty("text")
			If  sValue = Trim(sUpdatedBy) Then
					AddVerificationInfoToResult  "PASS " , "Actual value matched with expected."
					SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
					CaptureSnapshot()
			Else
					 iModuleFail = 1
					 AddVerificationInfoToResult  "Fail" , "Actual value does not matched with expected."
			End If

		End If
			
	CaptureSnapshot()

	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : UsedProdMenuSelect_fn
' 	Description : This function fetches row id for required rows
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function UsedProdMenuSelect_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [UsedProdMenuSelect$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	Browser("index:=0").Page("index:=0").Sync
	sAboutRecord = adoData( "AboutRecord")  & ""
	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""
	If instr(sLocateColValue,"Promotion")>0 Then
		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
	End If

	Index = adoData( "Index")  & ""
	If Index="" Then
		Index=0
	End If


	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Used Product/Service")

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
						DictionaryTempTest_G(sKey1)=sValue
					End If
			CaptureSnapshot()
			End If
			adoData.MoveNext
			Loop ''


	If sAboutRecord = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Used Product/Service").SiebMenu("Menu").SiebSelectMenu "About Record"
	End If
	

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : AgreementDateChange_fn
' 	Description : This function fetches row id for required rows
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function AgreementDateChange_fn()

'Dim adoData	  
'    strSQL = "SELECT * FROM [UsedProdMenuSelect$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	 strpromotion = 	DictionaryTest_G.Item("ProductName")


	'Flow
	Browser("index:=0").Page("index:=0").Sync

		mth1= DateAdd("m",1,Now)
		mth1 = month(mth1)
		mth1 =Left(MonthName(mth1),3)
	'	msgbox mth1
		
		yr= year(Now)
		'yrPlus2 = year(yrPlus2)
	'	msgbox yr1
		yrMinus2=DateAdd("yyyy",-2,Now)
		yrMinus2 = year(yrMinus2)


	'	yrPlus1= DateAdd("yyyy",1,Now)
		'yrPlus1 = year(yrPlus1)
	'	msgbox yr1
		yrMinus1=DateAdd("yyyy",-1,Now)
		yrMinus1 = year(yrMinus1)
'		msgbox yr2

		If  instr(strpromotion,"12" )>0Then
			PlannedStartDate =  "04" &"/"& mth1 &"/"&yrMinus1
			EndDate =  "04" &"/"& mth1 &"/"&yr
		ElseIf instr(strpromotion,"24" )>0 Then
			PlannedStartDate = "04" &"/"& mth1 &"/"&yrMinus2
			EndDate = "04" &"/"& mth1 &"/"&yr
		End If

		
			
	'	msgbox DateRequested
	
	SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Agreement Line Detail").SiebApplet("Agreement").SiebCalendar("Planned start").SetText PlannedStartDate
	AddVerificationInfoToResult  "Info" , "Planned Start date set as : "&PlannedStartDate
	SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Agreement Line Detail").SiebApplet("Agreement").SiebCalendar("End").SetText EndDate
	AddVerificationInfoToResult  "info" , "end Date Set as : "&EndDate
	SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Agreement Line Detail").SiebApplet("Agreement").SiebMenu("Menu").Select "ExecuteQuery"	


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : SelectAgreementRowID_fn
' 	Description : This function fetches row id for required rows
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function SelectAgreementRowID_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [UsedProdMenuSelect$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	Browser("index:=0").Page("index:=0").Sync

	
	sAboutRecord = adoData( "AboutRecord")  & ""
	sCheckAgreementID = adoData( "CheckAgreementID")  & ""

	Index = adoData( "Index")  & ""
	If Index="" Then
		Index=0
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account Agreement List View","L3"
	Browser("index:=0").Page("index:=0").Sync

	If sCheckAgreementID ="Y"  Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Agreements List View").SiebMenu("Menu").Select "ExecuteQuery"	
	End If

	'Flow
	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Agreements List View")

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

			If sAboutRecord = "Y" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Agreements List View").SiebMenu("Menu").SiebSelectMenu "About Record"
			End If
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : SelectConditionalChargesRow_fn
' 	Description : This function fetches row id for required rows
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function SelectConditionalChargesRow_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [UsedProdMenuSelect$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	Browser("index:=0").Page("index:=0").Sync

	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""
	If instr(sLocateColValue,"Promotion")>0 Then
		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
	End If

	Index = adoData( "Index")  & ""
	If Index="" Then
		Index=0
	End If

		SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebScreenViews("ScreenViews").Goto "FS Agree Conditional Charge View","L4"

		Browser("index:=0").Page("index:=0").Sync

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Renewal Escalator").SiebApplet("Line Items")
	'objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

	If sLocateCol <> "" Then
		res=LocateColumns (objApplet,sLocateCol,sLocateColValue,Index)
		If  Cstr(res)<>"True" Then
			 iModuleFail = 1
			AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
			Exit Function
		End If
	End If	

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : ConditionalChargesRowID_fn
' 	Description : This function fetches row id for required rows
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function ConditionalChargesRowID_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [UsedProdMenuSelect$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	Browser("index:=0").Page("index:=0").Sync
	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""
	If instr(sLocateColValue,"Promotion")>0 Then
		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
	End If

	Index = adoData( "Index")  & ""
	If Index="" Then
		Index=0
	End If

	Set objApplet =SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Renewal Escalator").SiebApplet("Conditional Charges")


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
	
		SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Renewal Escalator").SiebApplet("Conditional Charges").SiebMenu("Menu").SiebSelectMenu "About Record"
	
	

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : Contacts_LastName_fn
' 	Description : This function is used to used to click on contacts tab on accounts page and drilldown on last name.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Contacts_LastName_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ContactsLastName$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sOnlineAccountValidation = adoData( "OnlineAccountValidation")  & ""
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").Exist(5) Then
	Else
		AddVerificationInfoToResult "Error","Contacts tab is not clicked successfully. "
		iModuleFail = 1
		Exit Function
	End If
	
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").ActivateRow 0

	If sOnlineAccountValidation = "Y" Then
		sOnlineCheck = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").SiebCheckbox("Online ID").GetRoProperty("isenabled")
		If sOnlineCheck = "False" Then
			AddVerificationInfoToResult  "Info" , "As expected Online check box is disabled i.e not editable after saving account on Contacts tab page."
			CaptureSnapshot()
			Exit Function
		Else
			AddVerificationInfoToResult  "Error" , "Online check box is enabled i.e editable after saving account on Contacts tab page."
			iModuleFail = 1
			CaptureSnapshot()
			Exit Function
		End If

	End If

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts")
	Do while Not adoData.Eof

		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	

		UpdateSiebList objApplet,sUIName,sValue
		AddVerificationInfoToResult "Info","Last Name  is drilldown successfully. " 
		adoData.MoveNext
	Loop

		CaptureSnapshot()			
  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : Accounts_PrePost_fn
' 	Description : This function is used to used to click on contacts tab on accounts page and drilldown on last name.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Accounts_PrePost_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AccountsPrePost$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sClickNew = adoData( "ClickNew")  & ""

	'Flow
	If sClickNew = "Yes" Then
	SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebScreenViews("ScreenViews").Goto "Contact Detail - Accounts View","L3"
	End If

	If SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Companies").Exist(5) Then
	Else
		AddVerificationInfoToResult "Error","Accounts tab on Contacts page is not clicked successfully. "
		iModuleFail = 1
		Exit Function
	End If

	If sClickNew = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Companies").SiebButton("New").SiebClick False
		Exit Function
	End If
	
	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Companies")
	Do while Not adoData.Eof
		sLocateCol =  adoData("LocateCol") & ""
		sLocateColValue =  adoData("LocateColValue") & ""
		index = 0
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	
		If sLocateCol<>"" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,index)
		
			If res = False Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If	

		UpdateSiebList objApplet,sUIName,sValue
			If instr (sUIName, "CaptureTextValue1") > 0 AND instr (sUIName, "Account no.") > 0 Then
				If DictionaryTest_G.Exists("PrePostAccountNo") Then
					DictionaryTest_G.Item("PrePostAccountNo")=sValue
					AddVerificationInfoToResult  "Info" , "New Account no created is : " & DictionaryTest_G.Item("PrePostAccountNo")
				else
					DictionaryTest_G.add "PrePostAccountNo",sValue
				End If
			End If
		adoData.MoveNext
	Loop

	Browser("index:=0").Page("index:=0").Sync
'	DictionaryTest_G.Item("AccountNo") = "7000220161"

	If SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Pick Account").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").SiebSetText DictionaryTest_G.Item("AccountNo")
		SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
		wait 2
		CaptureSnapshot()			
		SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
	Else
		AddVerificationInfoToResult  "Error" , "Parent account open pop up is not clicked successfully."
		iModuleFail = 1
		Exit Function
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Contact Accounts").SiebApplet("Companies").SiebMenu("Menu").Select "ExecuteQuery"

	Browser("index:=0").Page("index:=0").Sync

  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : SearchAccountPrePost_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function SearchAccountPrePost_fn()

   Dim adoData	  
'    strSQL = "SELECT * FROM [SearchAccount$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	sAccountNumber = adoData( "Account_Number")  & ""
    sAccountNumber = DictionaryTest_G.Item("PrePostAccountNo")
'	sAccountName = adoData( "AccountName")  & ""

	If sAccountNumber = "" Then
		AddVerificationInfoToResult "Fail","Account number is null and not retrieved from data base"
		iModuleFail = 1
		Exit Function
	End If

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If	

	
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Query").SiebClick False
    
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Account no.").SiebSetText sAccountNumber
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Go").SiebClick False

	rowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").RowsCount
	
	If rowCount = 1 Then
		AddVerificationInfoToResult "Search Account - Pass","Account Search is done successfully."
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Search Account - Fail","Account Search is failed"
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End function


'#################################################################################################
' 	Function Name : Accounts_AddressLine
' 	Description : This function is used to used to click on contacts tab on accounts page and drilldown on last name.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Accounts_AddressLine()

	   call SetObjRepository ("Account",sProductDir & "Resources\")

		Browser("index:=0").Page("index:=0").Sync

		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").Exist(5) Then
		Else
			AddVerificationInfoToResult  "Error" , "Account Name is not drilled down successfully."
			iModuleFail = 1
			Exit Function
		End If

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup

'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Add Address").SiebList("List").SiebPicklist("Address status").Select "Validated"
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SiebSetText "E1W 2RL"

	   On error resume next
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Add Address").SiebList("List").SiebPicklist("Address status").ProcessKey "EnterKey"
		err.clear
		
'SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Add Address").SiebList("List").SiebPicklist("Address status").Select

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Addresses").SiebButton("Add").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Addresses").SiebButton("OK").SiebClick False
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Addresses_2").SiebButton("Add >").SiebClick False
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Addresses_2").SiebButton("OK").SiebClick False

		Browser("index:=0").Page("index:=0").Sync


End Function


'#################################################################################################
' 	Function Name : PrePost_Upgrade_fn
' 	Description : This function performs Tariff Migration,Pre/Post
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function PrePost_Upgrade_fn()



	'Get Data
	Dim sStartingWith

	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	sStartingWith = DictionaryTest_G.Item("ProductName")
	sAccountNumber = DictionaryTest_G.Item("PrePostAccountNo")
'	sAccountNumber ="7000259025"

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").Exist(5) Then
	Else
		AddVerificationInfoToResult "Error","Pre/Post is not selected successfully from UsedProduct Services view."
		iModuleFail = 1
		Exit Function
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebText("New Billing Account").OpenPopup
	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Account").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").SiebSetText sAccountNumber
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
	Else
		AddVerificationInfoToResult "Error","New Billing Account Open PopUp is no clicked successfully."
		iModuleFail = 1
		Exit Function	
	End If

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebText("New Billing Profile").OpenPopup

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing profile").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billing profile").SiebButton("OK").SiebClick False
	Else
		AddVerificationInfoToResult "Error","New Billing Profile Open PopUp is no clicked successfully."
		iModuleFail = 1
		Exit Function	
	End If

	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("classname:=SiebApplication").SetTimeOut(600)
	Recovery.Enabled=False
	On error resume next
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebText("Target Promotion").OpenPopup
	
	Browser("index:=0").Page("index:=0").Sync
'	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebPicklist("micclass:=SiebPicklist","repositoryname:=PopupQueryCombobox").Exist(200) Then
	
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebPicklist("micclass:=SiebPicklist","repositoryname:=PopupQueryCombobox").SiebSelect "Promotion Name"
	
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebText("micclass:=SiebText","repositoryname:=PopupQuerySrchspec").SiebSetText sStartingWith
	CaptureSnapshot()
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebButton("micclass:=SiebButton","repositoryname:=PopupQueryExecute").SiebClick False
	CaptureSnapshot()
'	Else 
'		AddVerificationInfoToResult "Error","Target Promotion  pop up did not occur in time"
'		iModuleFail = 1
'		Exit Function	
'End If
'	Browser("index:=0").Page("index:=0").Sync

'	rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebList("micclass:=SiebList","repositoryname:=SiebList").SiebListRowCount

'	If rowCnt <> 0 Then
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:=ISS Promotion"&".*").SiebButton("micclass:=SiebButton","repositoryname:=Execute").SiebClick False
'		
''		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("micclass:=SiebApplet","repositoryname:= .*" &"Pick Applet").SiebButton("micclass:=SiebButton","repositoryname:=HTML MiniButton2").SiebClick False
'	Else
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error","No rows displayed after clicking on Go button."
'		Exit Function
'	End If

	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebButton("Continue").Click
	Err.Clear
	Recovery.Enabled=True
	SiebApplication("classname:=SiebApplication").SetTimeOut(30)

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If DictionaryTest_G.Exists("AccountNo") Then
			DictionaryTest_G.Item("AccountNo")=sAccountNumber 
		else
			DictionaryTest_G.add "AccountNo",sAccountNumber 
		End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : BilledProductServices_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function BilledProductServices_fn()


	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [UsedProductServices$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	sLocateCol = adoData( "LocateCol")  & ""
	sLocateColValue = adoData( "LocateColValue")  & ""
	sUIName = adoData( "UIName")  & ""
	sValue = adoData( "Value")  & ""

	iIndex = adoData( "Index")  & ""
	If iIndex="" Then
		iIndex=0
	End If

'SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service").SiebList("List").DrillDownColumn

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billed Product/Service")
'	Call ShowMoreButtonUsedProductsServices_fn()
	If sLocateCol <> "" Then

		res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
		If  Cstr(res)<>"True" Then
			 iModuleFail = 1
			AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
			Exit Function
		End If
	End If	


	If sUIName <> "" Then
		If sValue = "ID" Then
			sValue = Trim(Replace(sValue,"ID",DictionaryTest_G.Item("AgreementId")))
		End If
		UpdateSiebList objApplet,sUIName,sValue
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
'               Function Name : OnlineAccountValidation_fn
'               Description : This function log into Siebel Application
'   Created By :  Tarun
'               Creation Date :        
'##################################################################################################
Public Function OnlineAccountValidation_fn()

	'Get Data
	Dim adoData        
    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	sTitle=adoData("Title")  & ""
	sAccount_Type=adoData("Account_Type")  & ""
	sAccount_SubCategory=adoData("AccountSubCategory")  & ""
	sAccount_Category=adoData("AccountCategory")  & ""
    sOnlineAccount=adoData("OnlineAccount")  & ""
  

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Browser("index:=0").Page("index:=0").Sync

	If sAccountTab<>"No" Then
	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If
		
	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "PageTabs not present"
		Exit Function
	End If
	End If

	'To verify if account no, account type and acount status are blank before clicking new button

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("New").SiebClick False

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account type").SiebSelect sAccount_Type

		If Ucase(sAccount_Category) <> "" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account Category").SiebSelect sAccount_Category
		End If
		If Ucase(sAccount_SubCategory) <> "" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account SubCategory").SiebSelect sAccount_SubCategory
		End If

		sOnlineCheck = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Create online account").GetRoProperty("ischecked")
		If sOnlineCheck = "False" Then
			AddVerificationInfoToResult  "Info" , "Create Online checkbox is disabled before selecting Title from dropdown."
			CaptureSnapshot()
		Else
			AddVerificationInfoToResult  "Error" , "Create Online checkbox is enabled before selecting Title from dropdown. It should be disabled."
			CaptureSnapshot()
			iModuleFail = 1
		End If
	
		 SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Title").SiebSelect sTitle

		sOnlineCheck = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCheckbox("Create online account").GetRoProperty("ischecked")
		If sOnlineCheck = "True" Then
			AddVerificationInfoToResult  "Info" , "Create Online checkbox is enabled after selecting Title from dropdown."
			CaptureSnapshot()
		Else
			AddVerificationInfoToResult  "Error" , "Create Online checkbox is disabled after selecting Title from dropdown. It should be enabled."
			iModuleFail = 1
		End If

   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : AllProductServicesValidation_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function AllProductServicesValidation_fn()

	'Get Data
	Dim adoData  
    strSQL = "SELECT * FROM [AllServicesValidation$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sServiceType=adoData("ServiceType")  & ""

	If sServiceType = "OwnedProduct" Then
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service")
		colCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Owned Product/Service").SiebList("List").ColumnsCount
	ElseIf sServiceType = "UsedProduct" Then
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Used Product/Service")
		colCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Used Product/Service").SiebList("List").ColumnsCount
	ElseIf sServiceType = "BilledProduct" Then
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billed Product/Service")
		colCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Billed Product/Service").SiebList("List").ColumnsCount
	End If

	objApplet.SiebMenu("Menu").Select "ColumnsDisplayed"
	If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Exist(5) Then
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Click
		Browser("index:=0").Page("index:=0").Sync
	Else
		AddVerificationInfoToResult  "Error" ,"Columns displayed value is not selected from Menu items.."
		iModuleFail = 1
	End If

	'Flow 
	For loopVar = 1 to colCnt
		sReposName = objApplet.SiebList("List").GetColumnRepositoryNameByIndex(loopVar)
		sUIName = objApplet.SiebList("List").GetColumnUIName(sReposName)
		If sUIName = "Start date" Then
			val1 = loopVar
		End If

		If sUIName = "Service end date" Then
			val2 = loopVar
		End If
	
	Next
	
	If Cint(val2) = Cint(val1) + 1 Then
		AddVerificationInfoToResult  "Info" ,"Start date is beside Service end date for " & sServiceType & " view."
		CaptureSnapshot()
	Else
		AddVerificationInfoToResult  "Error" ,"Start date is not beside Service end date for " & sServiceType & " view."
		iModuleFail = 1
	End If

		objApplet.SiebList("List").ActivateRow 0
		CaptureSnapshot()

	Do while Not adoData.Eof
		sLocateCol =  adoData("LocateCol") & ""
		sLocateColValue =  adoData("LocateColValue") & ""
		index = 0
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	
		If sLocateCol<>"" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,index)
		
			If res = False Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

		UpdateSiebList objApplet,sUIName,sValue
		AddVerificationInfoToResult  "Info" ,"User Account is present in list for " & sServiceType & " view."
		CaptureSnapshot()
		adoData.MoveNext
	Loop

	Browser("index:=0").Page("index:=0").Sync
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
'   Function Name : NonValidatedAddressValidation_fn
'   Description : This function log into Siebel Application
'   Created By :  Tarun
'   Creation Date :        
'##################################################################################################
Public Function NonValidatedAddressValidation_fn()

	'Get Data
	Dim adoData        
    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
                
    sEmail = adoData("Email")  & ""
	sPost_Code = adoData("Post_Code")  & ""
	sAccount_Type=adoData("Account_Type")  & ""
	DictionaryTest_G.Add "sAccount_Type", sAccount_Type
	sAccount_SubCategory=adoData("AccountSubCategory")  & ""
	sAccount_Category=adoData("AccountCategory")  & ""
    sOnlineAccount=adoData("OnlineAccount")  & ""
    sPin1 = adoData("Pin1") & ""
	sPin2 = adoData("Pin2")
	sPin3 = adoData("Pin3")
	sPin4 = adoData("Pin4")
   sAddressLine=adoData("AddressLine")  & ""
   sUserName = adoData("User_name")  & ""
   sPostcode=adoData("Postcode")  & ""
'   sHouseNumber=adoData("HouseNumber")  & ""
    sSave=adoData("Save")  & ""  
	sCareTeam=adoData("CareTeam")  & ""
	sTradingAs = adoData("TradingAs")  & ""
	sAnonymous = adoData("Anonymous")  & ""
	sPhoneNumber = adoData("PhoneNumber")
	sEmail1 = adoData("Email1")
	sEmail2 = adoData("Email2")
	sAccountTab = adoData("AccountTab") & ""
	sNew = adoData("New") & ""
	sSkipAddress = adoData("AddressLine") & "" ''Should be "N" in account sheet to skip entering address
	sLegal = adoData("Legal Status") & ""
	sAnonymousDrillDown = adoData("AnonymousDrillDown") & ""
	sGender = adoData("Gender") & ""
	sPriorityFaultRepair = adoData("PriorityFaultRepair") & ""
    sMatchCompany = adoData("MatchCompany") & ""
	sRegistrationNo = adoData("RegistrationNo") & ""
	sFirst_Name = adoData( "First_Name")  & ""
	sLast_Name = adoData("Last_Name")  & ""
	sTitle=adoData("Title")  & ""
	sBirth_Date = adoData("Birth_Date")  & ""
	sQASPostCode = adoData("QASPostCode") & ""''Enter QAS postcode in the respective column

	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If	

	Browser("index:=0").Page("index:=0").Sync


	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
	End If
		
	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "PageTabs not present"
		Exit Function
	End If


	'To verify if account no, account type and acount status are blank before clicking new button

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("New").SiebClick False

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account type").SiebSelect sAccount_Type


		If Ucase(sAccount_Category) <> "" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account Category").SiebSelect sAccount_Category
		End If
		If Ucase(sAccount_SubCategory) <> "" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Account SubCategory").SiebSelect sAccount_SubCategory
		End If
	
    SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebPicklist("Title").SiebSelect sTitle
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("First name").SiebSetText sFirst_Name
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Last name").SiebSetText sLast_Name
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebCalendar("Date of birth").SiebSetText sBirth_Date

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebButton("Update Email").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebText("Email").SiebSetText sEmail1
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebPicklist("SiebPicklist").SetText sEmail2
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebText("Re-type Email").SiebSetText sEmail1
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebPicklist("SiebPicklist_2").SetText sEmail2
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Siebel").SiebButton("OK").SiebClick False

     SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN1").SiebSetText sPin1
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN2").SiebSetText sPin2
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN3").SiebSetText sPin3
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("PIN4").SiebSetText sPin4

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Details").SiebText("Address Line 1").OpenPopup
    Browser("index:=0").Page("index:=0").Sync

	If sMatchCompany = "Y" Then
		sPostCode = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").GetROProperty("text")
		If sPostCode = "" Then
			AddVerificationInfoToResult  "Info" ,"Addresss applet is opened in query mode while creating new account."
			CaptureSnapshot()
			Exit Function
		Else
			AddVerificationInfoToResult  "Error" ,"Addresss applet is not opened in query mode."
			iModuleFail = 1
			Exit Function
		End If
	End If
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").SiebSetText sPostcode
   On error resume next
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").ProcessKey "EnterKey"
	err.clear

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts").SiebApplet("Account Addresses").SiebButton("Add >").SiebClick sPopUp

    CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : AddCreditBalance_fn
' 	Description : This function is used to click on Real Time Balance
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function AddCreditBalance_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AddCreditBalance$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
		sReason = adoData( "Reason")  & ""
		sAmount = adoData( "Amount")  & ""
		sLocate = adoData( "Locate")  & ""
	
		sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If	

	
	'sOR
			call SetObjRepository ("Account",sProductDir & "Resources\")

				SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebScreenViews("ScreenViews").Goto "CMU Adjustment History View","L3"

				rowNum = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("Product / Services").SiebList("List").SiebListRowCount

				If rowNum = 0 Then
					AddVerificationInfoToResult  "Fail" , "Adjustment Tab is not clicked successfully."
					iModuleFail = 1
					Exit Function
				End If
			
				res = False
				For loopVar = 0 to rowNum-1
				strVal = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("Product / Services").SiebList("List").SiebGetCellText("Name",loopVar)

						If Instr(sLocate,strVal) > 0 Then
							SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("Product / Services").SiebList("List").ActivateRow loopVar
							res = True
							Exit For
						End If
				Next
			

					SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("Product / Services").SiebButton("Credit Funds").SiebClick False
					Browser("index:=0").Page("index:=0").Sync

			If  SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("SiebApplet").Exist(5)Then
				SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("SiebApplet").SiebPicklist("Adjustment reason").SiebSelect sReason
				SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("SiebApplet").SiebCalculator("Authorisation amount").SetText sAmount
				SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("SiebApplet").SiebButton("Submit").SiebClick sPopUp
			End If
	
			rowNum = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("Adjustment Activities").SiebList("List").SiebListRowCount
'			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("Adjustment Activities").SiebList("List").DrillDownColumn

				If rowNum = 0 Then
					AddVerificationInfoToResult  "Fail" , "Credit Funds not added successfully."
					iModuleFail = 1
					Exit Function
				Else
						Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Adjustments").SiebApplet("Adjustment Activities")
					Do while Not adoData.Eof
							sLocateCol = adoData( "LocateCol")  & ""
							sLocateColValue = adoData( "LocateColValue")  & ""
							sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
'							sValue = adoData( "Value")  & ""
'							sUIName = adoData( "UIName")  & ""
							'sExist = adoData( "Exist")  & ""
						If sLocateCol <> "" Then
							res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
							
'
						End If
						CaptureSnapshot()
						AddVerificationInfoToResult  "Pass" , "Credit Funds is added successfully."
						adoData.MoveNext
			Loop
				End If
			
				
				If Err.Number <> 0 Then
						iModuleFail = 1
						AddVerificationInfoToResult  "Error" , Err.Description
				End If
	
End Function


'#################################################################################################
' 	Function Name : MainScreenServiceRequest_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By : Tarun
'	Creation Date :        
'##################################################################################################
Public Function MainScreenServiceRequest_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""
	sOwnerValidation = adoData( "OwnerValidation")  & ""
	sChangeOwner = adoData( "ChangeOwner")  & ""
	sOwnerLogin = adoData( "OwnerLogin")  & ""
	sServiceIdValidation = adoData( "ServiceIdValidation")  & ""

	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	
	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Service Request Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Service Request Screen"
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoView "Personal Service Request List View"
	End If


	Browser("index:=0").Page("index:=0").Sync

    	If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").Exist(5) Then
		AddVerificationInfoToResult  "Info" , "Service Request screen is displayed successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Service Request screen is not displayed successfully."
		iModuleFail = 1
		Exit Function
	End If

	If sServiceIdValidation = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebMenu("Menu").Select "NewQuery"
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").SiebText("ID").SiebSetText DictionaryTest_G.Item("ServiceId")
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebButton("Go").SiebClick False
		CaptureSnapshot()
	End If

	If sOwnerValidation = "Y" Then
		sServiceId = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").SiebText("ID").GetRoProperty("text")
		DictionaryTest_G.Item("ServiceId") = sServiceId
		sCreatedBy = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").SiebText("Created by").GetRoProperty("text")
		sOwner = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").SiebText("Owner").GetRoProperty("text")

		If  sServiceIdValidation = "Y" Then
			rwCnt = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").RowsCount
			If rwCnt = "1" Then
				AddVerificationInfoToResult  "Info" , "Service Request is created successfully when Created by is  " & sCreatedBy & " Owner is  " & sOwner
			Else
				AddVerificationInfoToResult  "Error" , "Service Request is not  created successfully when Created by is  " & sCreatedBy & " Owner is  " & sOwner
				iModuleFail = 1
			End If
		End If

	End If

	If Ucase(sClickNew) = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebButton("New").SiebClick sPopUp
		If sOwnerValidation = "Y" Then
			If sCreatedBy = sOwner Then
				AddVerificationInfoToResult  "Info" , "Error message is disaplyed successfully when both Created by account - " & sCreatedBy & " and Owner account - " & sOwner & " acccount are same and Service Id created is : " & DictionaryTest_G.Item("ServiceId")
			End If
		End If
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").ActivateRow 0

	Set obj = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests")


	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	

		If sValue = "NOW" Then
			TimeRequested = FormatDateTime (Time, vbShortTime)& Right (Time, 6)
			MN =Left( MonthName(Month(Date)),3)
			DateRequested = Date					
			DateRequested = day(DateRequested+1) & "/" & MN & "/" &  year(DateRequested) &" "& TimeRequested
			sValue = DateRequested								
		End If

		If sUIName <> "" Then
			UpdateSiebList obj,sUIName,sValue
		End If

		If instr (sUIName, "Contact") > 0 Then
			If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Contact").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Contact").SiebButton("OK").SiebClick False
			Else
				AddVerificationInfoToResult  "Error" , "Accounts Open pop up is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If
		End If
	
		If instr (sUIName, "Account name") > 0 Then
			If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Account").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				CaptureSnapshot()
			Else
				AddVerificationInfoToResult  "Error" , "Account name Open pop up is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If
		End If
		adoData.MoveNext
	Loop

	If sChangeOwner = "Y" Then
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").SiebText("Owner").OpenPopup
		If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Service Request Owner").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Service Request Owner").SiebPicklist("SiebPicklist").SiebSelect "User ID"
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Service Request Owner").SiebText("SiebText").SiebSetText sOwnerLogin
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Pick Service Request Owner").SiebButton("Go").SiebClick False
		Else
			AddVerificationInfoToResult  "Error" , "Owner Open pop up is not clicked successfully."
			iModuleFail = 1
		End If
	
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : VerifyAttributesTab_fn
' 	Description : This function validates required fields in the Attributes tab under Product/Services
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################
Public Function VerifyAttributesTab_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ProductServices$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Attributes")
			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

			Do while Not adoData.Eof
'				sUIName = adoData( "UIName")  & ""
'				sValue = adoData( "Value")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""
				
'					iIndex = adoData( "Index")  & ""
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
'				If sUIName <> "" Then
'					 UpdateSiebList objApplet,sUIName,sValue
'					CaptureSnapshot()
'				End If
			adoData.MoveNext
			Loop ''

		If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : CustomerCommsServiceRequest_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CustomerCommsServiceRequest_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [CustomerCommsServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""
	sCategory = adoData( "Category")  & ""
	sSubCategory = adoData( "SubCategory")  & ""
	sResolution = adoData( "Resolution")  & ""
	sType = adoData( "Type")  & ""
	sArea = adoData( "Area")  & ""
	sSubArea = adoData( "SubArea")  & ""
	sOwnerLogin = adoData( "OwnerLogin")  & ""
	sDPAValidation = adoData( "DPAValidation")  & ""
    


	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")


	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Activities Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
	End If

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").Exist(5) Then
		AddVerificationInfoToResult  "Info" , "Customer  Comms screen is displayed successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Customer Comms screen is not displayed successfully."
		iModuleFail = 1
		Exit Function
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebButton("New").SiebClick  False

	SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").ActivateRow 0
	CaptureSnapshot()

	Set obj = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms")


	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	

		If sUIName <> "" Then
			UpdateSiebList obj,sUIName,sValue
		End If

		adoData.MoveNext
	Loop

	If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").Exist(5)  Then
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Category").SiebSelect sCategory
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Sub-category").SiebSelect sSubCategory
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("Resolution").SiebSelect	sResolution
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("DPA validation").SiebSelect	sDPAValidation
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebButton("Create SR").SiebClick False
		AddVerificationInfoToResult  "Info" , "SR is created successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Drill down is not done successfully at customer comms screen."
		iModuleFail = 1
		Exit Function
	End If

	If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Type").SiebSelect sType
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Area").SiebSelect sArea
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Sub area").SiebSelect sSubArea
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebTextArea("Notes<img src=\'images/icon_re").SetText "abc"
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebTextArea("Request").SetText "abc"

			TimeRequested = FormatDateTime (Time, vbShortTime)& Right (Time, 6)
			MN =Left( MonthName(Month(Date)),3)
			DateRequested = Date					
			DateRequested = day(DateRequested+1) & "/" & MN & "/" &  year(DateRequested) &" "& TimeRequested
			sSLADate = DateRequested	
			
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebCalendar("SLA date").SetText  sSLADate
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebText("Contact").OpenPopup

			If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Pick Contact").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Pick Contact").SiebButton("OK").SiebClick False
			Else
				AddVerificationInfoToResult  "Error" , "Contact Open pop up is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If

			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebText("Account name").OpenPopup

			If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Pick Account").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
			Else
				AddVerificationInfoToResult  "Error" , "Account name Open pop up is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If

			AddVerificationInfoToResult  "Info" , "Service Request details are filled successfully."

			CaptureSnapshot()

			sServiceId = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebText("ID").GetRoProperty("text")
			DictionaryTest_G.Item("ServiceId") = sServiceId
			sOwner = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebText("Owner").GetRoProperty("text")
			sCreatedBy = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebText("Created by").GetRoProperty("text")
	
			If Instr(sPopUp,"You cannot raise a manager") > 0 Then
				SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("Submit SR").SiebClick False
				SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebMenu("Menu").Select "WriteRecord"
					If Browser("Siebel Call Center").Dialog("Siebel").Static("You cannot raise a manager_2").Exist(5) Then
						Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
						AddVerificationInfoToResult  "Info" , "Error message id displayed successfuly when both Created by - "  & sCreatedBy & " and Owner - " & sOwner & "  user login are same and Service Id created is - " &  DictionaryTest_G.Item("ServiceId")
					Else
						AddVerificationInfoToResult  "Error" , "Error message id is not displayed successfuly when both Created by - "  & sCreatedBy & " and Owner - " & sOwner & "  user login are same and Service Id created is - " &  DictionaryTest_G.Item("ServiceId")
						iModuleFail = 1
					End If				
	
				Browser("index:=0").Page("index:=0").Sync
	
				SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebText("Owner").OpenPopup
				If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Pick Service Request Owner").Exist(5) Then
					SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Pick Service Request Owner").SiebText("SiebText").SiebSetText sOwnerLogin
					SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Pick Service Request Owner").SiebButton("Go").SiebClick False
					AddVerificationInfoToResult  "Info" , "Owner User login is changed to " & sOwnerLogin
				Else
					AddVerificationInfoToResult  "Error" , "Owner Open poup is not clicked successfully."
					iModuleFail = 1
					Exit Function
				End If
			End If
	Else
		AddVerificationInfoToResult  "Error" , "Create SR button is not clicked successully."
		iModuleFail = 1
		Exit Function
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : Accounts_ServiceRequest_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By : Tarun
'	Creation Date :        
'##################################################################################################
Public Function Accounts_ServiceRequest_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""
	

	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").getROProperty("activeview") <> "Service Account List View" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Service Account List View","L3"
	End If

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Service Requests").Exist(5) Then
		AddVerificationInfoToResult  "Info" , "Service Request screen is displayed successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Service Request screen is not displayed successfully."
		iModuleFail = 1
		Exit Function
	End If


	If Ucase(sClickNew) = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Service Requests").SiebButton("New").SiebClick sPopUp
	End If

	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Service Requests").SiebList("List").ActivateRow 0

	Set obj = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Service Requests")


	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	
		sChangeOwner = adoData( "ChangeOwner")  & ""
		sOwnerValidation = adoData( "OwnerValidation")  & ""
		sOwnerLogin = adoData( "OwnerLogin")  & ""

		If sUIName <> "" Then
			UpdateSiebList obj,sUIName,sValue
		End If

		If instr(sUIName,"CaptureTextValue") > 0 Then
			DictionaryTest_G.Item("ServiceId") = sValue
		End If

		If sOwnerValidation = "Y" Then
			AddVerificationInfoToResult  "Info" , "Error message is displayed successfully when both Created by and Owner logins are same and Service Id created is - " & DictionaryTest_G.Item("ServiceId")
		End If

		If sChangeOwner = "Y" Then
			CaptureSnapshot()
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Pick Service Request Owner").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Pick Service Request Owner").SiebPicklist("SiebPicklist").SiebSelect "User ID"
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Pick Service Request Owner").SiebText("SiebText").SiebSetText sOwnerLogin
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Service Requests").SiebApplet("Pick Service Request Owner").SiebButton("Go").SiebClick False
			Else
				AddVerificationInfoToResult  "Error" , "Owner Open pop up is not clicked successfully."
				iModuleFail = 1
			End If
		End If
		adoData.MoveNext
	Loop


	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : AccountSummary_ServiceRequest_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By : Tarun
'	Creation Date :        
'##################################################################################################
Public Function AccountSummary_ServiceRequest_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""


	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").Exist(5) Then
		AddVerificationInfoToResult  "Info" , "Service Request screen is displayed successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Service Request screen is not displayed successfully."
		iModuleFail = 1
		Exit Function
	End If


	If Ucase(sClickNew) = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").SiebButton("New").SiebClick sPopUp
	End If

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests").SiebList("List").ActivateRow 0

	Set obj = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Service Requests")


	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	
		sChangeOwner = adoData( "ChangeOwner")  & ""
		sOwnerValidation = adoData( "OwnerValidation")  & ""
		sOwnerLogin = adoData( "OwnerLogin")  & ""

		If sUIName <> "" Then
			UpdateSiebList obj,sUIName,sValue
		End If

		If instr(sUIName,"CaptureTextValue") > 0 Then
			DictionaryTest_G.Item("ServiceId") = sValue
		End If

		If sOwnerValidation = "Y" Then
			AddVerificationInfoToResult  "Info" , "Error message is displayed successfully when both Created by and Owner logins are same and Service Id created is - " & DictionaryTest_G.Item("ServiceId")
		End If

		If sChangeOwner = "Y" Then
			CaptureSnapshot()
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebPicklist("SiebPicklist").SiebSelect "User ID"
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebText("SiebText").SiebSetText sOwnerLogin
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Pre to Post").SiebButton("Go").SiebClick False
				CaptureSnapshot()
			Else
				AddVerificationInfoToResult  "Error" , "Owner Open pop up is not clicked successfully."
				iModuleFail = 1
			End If
		End If
		adoData.MoveNext
	Loop


	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : ServiceRequestColumnValidation_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################

Public Function ServiceRequestColumnValidation_fn()

	'Get Data
	Dim adoData  
    strSQL = "SELECT * FROM [AllServicesValidation$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sServiceType =  adoData("ServiceType") & ""

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Service Request Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Service Request Screen"
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebScreenViews("ScreenViews").Goto "Personal Service Request List View","L2"
	End If


	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").Exist(5) Then
		AddVerificationInfoToResult  "Info" , "Service Request screen is displayed successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Service Request screen is not displayed successfully."
		iModuleFail = 1
		Exit Function
	End If

	If  sServiceType = "SRClosedDate" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebMenu("Menu").Select "NewQuery"
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").SiebPicklist("Status").Select "Closed"
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebButton("Go").SiebClick False
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").ActivateRow 0
			sSRClosedDate = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").SiebCalendar("Closed Date").GetRoProperty("isenabled")
			If sSRClosedDate = "False" Then
				AddVerificationInfoToResult  "Info" , "SR closed date is not editable when SR status is closed."
				CaptureSnapshot()
				Exit Function
			Else
				AddVerificationInfoToResult  "Error" , "SR closed date is editable when SR status is closed."
				iModuleFail = 1
				Exit Function
			End If
			
	End If

	If sServiceType = "ColumnValidation" Then
		
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebMenu("Menu").Select "ColumnsDisplayed"
		
		If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Exist(5) Then
			Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Click
			Browser("index:=0").Page("index:=0").Sync
		Else
			AddVerificationInfoToResult  "Error" ,"Columns displayed value is not selected from Menu items."
			iModuleFail = 1
			Exit Function
		End If
	
			Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests")
			colCnt = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests").SiebList("List").ColumnsCount
	
		'Flow 
		For loopVar = 1 to colCnt
			sReposName = objApplet.SiebList("List").GetColumnRepositoryNameByIndex(loopVar)
			sUIName = objApplet.SiebList("List").GetColumnUIName(sReposName)
			If sUIName = "Created By" Then
				val1 = loopVar
			End If
	
			If sUIName = "Owner" Then
				val2 = loopVar
			End If
		
		Next
		
		If Cint(val2) = Cint(val1) + 1 Then
			AddVerificationInfoToResult  "Info" ,"Created by cloumn exits at the left of Owner column"			
		Else
			AddVerificationInfoToResult  "Error" ,"Created by cloumn doesnot exists at the left of Owner column"
			iModuleFail = 1
		End If
	End If

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests")

		objApplet.SiebList("List").ActivateRow 0

	Do while Not adoData.Eof
		sLocateCol =  adoData("LocateCol") & ""
		sLocateColValue =  adoData("LocateColValue") & ""
		index = 0
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	
		If sLocateCol<>"" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,index)
		
			If res = False Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

		UpdateSiebList objApplet,sUIName,sValue
		adoData.MoveNext
	Loop

	Browser("index:=0").Page("index:=0").Sync
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : CustomerSummary_fn
' 	Description : This Function validates the required fields in the Customer Summary page
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function CustomerSummary_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [CustomerSummary$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)


    'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sValidation = adoData( "Validation")  & ""
    sView = adoData( "View")  & ""
    sClickOnOwnedProduct = adoData( "ClickOnOwnedProduct")  & ""

	    If  (sClickOnOwnedProduct = "Yes")Then

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Owned Product and Service view","L3"
           CaptureSnapshot()
		   AddVerificationInfoToResult  "PASS" ,"Customer summary tab has been clicked"
		   SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Owned Product and Service view","L4"
			CaptureSnapshot()
			AddVerificationInfoToResult  "PASS" ,"Owned product service has been displayed successfully"
			Exit function

		End If


			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Owned Product and Service view","L3"

		    
			If sValidation = "Y" Then
			
						If sView = "Customer Summary" Then
							Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience")
							colCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience").SiebList("List").ColumnsCount
						
						End If

							If sView = "User Account" Then
							Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience")
							colCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience").SiebList("List").ColumnsCount

							For loopVar = 1 to colCnt
									sReposName = objApplet.SiebList("List").GetColumnRepositoryNameByIndex(loopVar)
									sUIName = objApplet.SiebList("List").GetColumnUIName(sReposName)
									If sUIName = "User Account" Then
										val1 = loopVar
									End If
							
									If sUIName = "PaymentType" Then
										val2 = loopVar
									End If
								
								Next
								
								If Cint(val2) = Cint(val1) + 2 Then
									AddVerificationInfoToResult  "Info" ,"User Account is before  Payment Type  for " & sView & " view."
									CaptureSnapshot()
								Else
									AddVerificationInfoToResult  "Error" ,"User Account is not before  Payment Type  for " & sView & " view."
									iModuleFail = 1
								End If
							
									objApplet.SiebList("List").ActivateRow 0
									CaptureSnapshot()		
						End If
					
'						objApplet.SiebMenu("Menu").Select "ColumnsDisplayed"
'						If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Exist(5) Then
'							Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Click
'							Browser("index:=0").Page("index:=0").Sync
'						Else
'							AddVerificationInfoToResult  "Error" ,"Columns displayed value is not selected from Menu items.."
'							iModuleFail = 1
'						End If
			
			
			
						Do while Not adoData.Eof
							sLocateCol =  adoData("LocateCol") & ""
							sLocateColValue =  adoData("LocateColValue") & ""
							index = 0
							sUIName = adoData( "UIName")  & ""
							sValue = adoData( "Value")  & ""	
							If sLocateCol<>"" Then
					
								res=LocateColumns (objApplet,sLocateCol,sLocateColValue,index)
							
								If res = False Then
									iModuleFail = 1
									AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
									Exit Function
								End If
							End If
					
							UpdateSiebList objApplet,sUIName,sValue
							If  instr (sUIName,"Diff. Payer")>0 Then
								AddVerificationInfoToResult  "Info" ,"Diff. Payer or User is present in list for " & sView & " view."
							Else If  instr (sUIName,"User Account")>0 Then
									AddVerificationInfoToResult  "Info" ,"User account is present in list for " & sView & " view."
								End If
							End If
							
							CaptureSnapshot()
							adoData.MoveNext
						Loop
			End If		
			Browser("index:=0").Page("index:=0").Sync
		   
			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If
	
	
			   
'     Else 
'   SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Billed Product and Service view","L4"
'

    Browser("index:=0").Page("index:=0").Sync
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

''#################################################################################################
' 	Function Name : ValidateAgreementStatus_fn
' 	Description :  This function Validate the status of  Agreement ID.
'   Created By :   Pushkar
'	Creation Date :    30/03/2017
''##################################################################################################
Public Function ValidateAgreementStatus_fn()


	Dim rwNum
	Dim loopVar
	Dim strProduct
	Dim sProduct
	Dim adoData
	Dim strSQL

 
    strSQL = "SELECT * FROM [AgreementStatus$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	Browser("index:=0").Page("index:=0").Sync

  Set objApplet =SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Agreement Line Detail").SiebApplet("Line Items")

  Do while Not adoData.Eof
'				sUIName = adoData( "UIName")  & ""
'				sValue = adoData( "Value")  & ""
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
'				If sUIName <> "" Then
'					 UpdateSiebList objApplet,sUIName,sValue
'					CaptureSnapshot()
'				End If
			adoData.MoveNext
			Loop ''

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function
'#################################################################################################
' 	Function Name : SearchAccountSADMIN_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Zeba
'	Creation Date :        
'##################################################################################################
Public Function SearchAccountSADMIN_fn()

		   Dim adoData	  
		'    strSQL = "SELECT * FROM [SearchAccount$] WHERE  [RowNo]=" & iRowNo
		'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
		'	
		'	'sOR
			call SetObjRepository ("Account",sProductDir & "Resources\")
		
			sAccountNumber = DictionaryTest_G.Item("AccountNo")
		
			If sAccountNumber = "" Then
				AddVerificationInfoToResult "Fail","Account number is null and not retrieved from data base"
				iModuleFail = 1
				Exit Function
			End If
		
			Browser("index:=0").Page("index:=0").Sync
		
			SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
		
			If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
				SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
			End If
		
			If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Accounts Screen" Then
				SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Accounts Screen"
			End If	
		
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "All Accounts across My Organizations","L2"
		
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts Across My").SiebApplet("Account Details").SiebButton("Query").SiebClick False
			
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts Across My").SiebApplet("Account Details").SiebText("Account no.").SiebSetText sAccountNumber
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts Across My").SiebApplet("Account Details").SiebButton("Go").SiebClick False
		
			If sAccountName <> "" Then
				sAccountNames = Trim(SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts Across My").SiebApplet("Account Details").SiebText("Account name").SiebGetRoProperty("text"))
		
					If DictionaryTest_G.Exists("AccountName") Then
						DictionaryTest_G.Item("AccountName")= sAccountNames
					Else
						DictionaryTest_G.add "AccountName",sAccountNames
					End If
			End If
		
			rowCount = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts Across My").SiebApplet("Accounts").SiebList("List").RowsCount
		
			If rowCount = 1 Then
				AddVerificationInfoToResult "Search Account - Pass","Account Search is done successfully."
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("All Accounts Across My").SiebApplet("Accounts").SiebList("List").SiebDrillDownColumn "Account name",0
			Else
				iModuleFail = 1
				AddVerificationInfoToResult "Search Account - Fail","Account Search is failed"
			End If
		
			Browser("index:=0").Page("index:=0").Sync
		
			CaptureSnapshot()
		
			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If

End function


'#############################################################################################################################
' 	Function Name : GotoAddresses_fn()
' 	Description : This function is used to click on Profile tab on Account page and click on the new buttom and verifies the applet is open in query mode.
'   Created By :  Vishwa
'	Creation Date :        
'#############################################################################################################################
Public Function GotoAddresses_fn()

  'Dim adoData	  
		'    strSQL = "SELECT * FROM [SearchAccount$] WHERE  [RowNo]=" & iRowNo
		'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
		'	
		'	'sOR
			call SetObjRepository ("Account",sProductDir & "Resources\")

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Business Address View","L3"
			Browser("index:=0").Page("index:=0").Sync
			CaptureSnapshot()

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").SiebButton("New").SiebClick False


		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Add Address").SiebList("List").Exist(5) then
			sPostCode = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Add Address").SiebList("List").SiebText("Postcode").GetROProperty("text")
			If sPostCode = "" Then
				AddVerificationInfoToResult  "Info" ,"Addresss applet is opened in query mode."
				CaptureSnapshot()
				Exit Function
			Else
				AddVerificationInfoToResult  "Error" ,"Addresss applet is not opened in query mode."
				iModuleFail = 1
				Exit Function
			End If
		Else
			AddVerificationInfoToResult  "Error" ,"New Button is not clicked successfully."
			iModuleFail  = 1
		End If
		
End function
'#################################################################################################
' 	Function Name : VerifyCustCommsList_fn()
' 	Description : This function is used to verify required fields on Customer comms list screen
'   Created By :  Zeba
'	Creation Date :        
'##################################################################################################

Public Function VerifyCustCommsList_fn()

   Dim adoData	  
	strSQL = "SELECT * FROM [CreateNewCustComms$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

sAccountNumber = adoData( "Account_Number")  & ""
''DictionaryTest_G.Item("AccountNo") ="7000215515"
    sAccountNumber = DictionaryTest_G.Item("AccountNo")
'	sAccountName = adoData( "AccountName")  & ""

	sDPAValidationInLineItems = adoData("DPAValidationInLineItems")  & ""
	sGotoScreen = adoData("GotoScreen")  & ""
	sValidateAccNo = adoData("ValidateAccNo")  & ""
	sViews = adoData("Views")  & ""
	sDPA_validation = adoData("DPA_validation") & ""

	sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If
'	sAccountNumber = "7000263611"

	If sAccountNumber = "" Then
		AddVerificationInfoToResult "Fail","Account number is null and not retrieved from data base"
		iModuleFail = 1
		Exit Function
	End If

			If sGotoScreen = "Y" Then
			
				SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
				Browser("index:=0").Page("index:=0").Sync
			
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Activity List View","L2"
			
			
				CaptureSnapshot()
			
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebButton("New").SiebClick False
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebText("Account No").OpenPopup
				Browser("index:=0").Page("index:=0").Sync
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").SiebSetText sAccountNumber
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").ProcessKey "EnterKey"
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				CaptureSnapshot()
				
			
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebText("Contact").OpenPopup
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Contact").SiebPicklist("SiebPicklist").SiebSelect "Account no."
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Contact").SiebText("SiebText").SiebSetText sAccountNumber
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Contact").SiebButton("Go").Click
			
			
		End If

	If sValidateAccNo = "Y" Then
			'On error resume Next
			SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
			'err.clear
			Browser("index:=0").Page("index:=0").Sync
						If sViews = "MyCustComms" Then
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Activity List View","L2"
								CaptureSnapshot()
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebButton("New").SiebClick False
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebText("Account No").OpenPopup
		'																																																																																																		
								Browser("index:=0").Page("index:=0").Sync
								CaptureSnapshot()
								Recovery.Enabled=False
								On error resume Next
									SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").ProcessKey "EnterKey"
								err.clear
								Recovery.Enabled=True
'					SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Activity").SiebButton("New").Click

										If Browser("Siebel Call Center").Dialog("Message from webpage").Static("Please ensure you have").Exist(4) then
												CaptureSnapshot()
												Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
												AddVerificationInfoToResult "PASS" , "Popup came as expected"
					'							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebText("Account No").Click
												SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebButton("Cancel").Click
												CaptureSnapshot()
												Else
												iModuleFail = 1
												AddVerificationInfoToResult "FAIL" , "Popup did not occured"
												Exit Function
										End If
			
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebText("Account No").OpenPopup
								Browser("index:=0").Page("index:=0").Sync
								
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").SiebSetText sAccountNumber
								CaptureSnapshot()
								On error resume Next
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").ProcessKey "EnterKey"
								err.clear
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
								CaptureSnapshot()
'								Exit Function
                    
				ElseIf sViews = "AllCustComms" Then
							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "VF All Activity List View Across My Organization","L2"
							CaptureSnapshot()
							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Activity").SiebButton("New").Click
							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Activities").SiebList("List").SiebText("Account No").OpenPopup
							Browser("index:=0").Page("index:=0").Sync
							CaptureSnapshot()
							Recovery.Enabled=False
							On error resume Next
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").ProcessKey "EnterKey"
							err.clear
							Recovery.Enabled=True
							If Browser("Siebel Call Center_2").Dialog("Message from webpage").Static("Please ensure you have").Exist(4) then
												CaptureSnapshot()
												Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Click
												AddVerificationInfoToResult "PASS" , "Popup came as expected"
					'							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebText("Account No").Click
												SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Pick Account").SiebButton("Cancel").Click
												CaptureSnapshot()
							Else
												iModuleFail = 1
												AddVerificationInfoToResult "FAIL" , "Popup did not occured"
												Exit Function
							End If

							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Activities").SiebList("List").SiebText("Account No").OpenPopup
							Browser("index:=0").Page("index:=0").Sync
								
							SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").SiebSetText sAccountNumber
							CaptureSnapshot()
							On error resume Next
								SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Pick Account").SiebList("List").SiebText("Account no.").ProcessKey "EnterKey"
                            	SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
								CaptureSnapshot()
							err.clear
							Exit Function

				End If
				
	End If

	If (sDPA_validation="NoContact") AND  (sDPAValidationInLineItems="Passed") Then

		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
		Browser("index:=0").Page("index:=0").Sync
	
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Activity List View","L2"
       CaptureSnapshot()

        SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebButton("New").SiebClick False
       CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebPicklist("DPA validation").SiebSelect sDPAValidationInLineItems
		CaptureSnapshot()
               If Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("Set/Reset Password").Exist Then
			CaptureSnapshot()
				AddVerificationInfoToResult "PASS" , "Set/Reset Password is Disabled as expected"
		Else
				iModuleFail = 1
				AddVerificationInfoToResult "FAIL" , "Set/Reset Password is Enabled"
				CaptureSnapshot()
				Exit Function
		End If
				If Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("Set/Reset PIN").Exist Then
					CaptureSnapshot()
				AddVerificationInfoToResult "PASS" , "Set/Reset PIN is Disabled as expected"
		Else
				iModuleFail = 1	
				AddVerificationInfoToResult "FAIL" , "Set/Reset PIN is Enabled"
				CaptureSnapshot()
				Exit Function
		End If

		If Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("Set/Reset Word and Hint").Exist Then
					CaptureSnapshot()
				AddVerificationInfoToResult "PASS" , "Set/Reset Word and Hint Disabled as expected"
		Else
				iModuleFail = 1
				AddVerificationInfoToResult "FAIL" , "Set/Reset Word and Hint is Enabled"
				CaptureSnapshot()
				Exit Function
		  End if

	End If


	If (sDPAValidationInLineItems="Passed" AND sDPA_validation <>"NoContact") Then

		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebPicklist("DPA validation").SiebSelect sDPAValidationInLineItems
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebMenu("Menu").Select "ExecuteQuery"

					If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebButton("Set/Reset Password").isenabled =True then
						CaptureSnapshot()
						AddVerificationInfoToResult "PASS" , "Set/Reset Password is Displayed successfully as expected"
					Else
						iModuleFail = 1
						AddVerificationInfoToResult "FAIL" , "Set/Reset Password is Not  Displayed successfully"
						Exit Function
					End If

					If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebButton("Set/Reset PIN").isenabled =True Then
						CaptureSnapshot()
						AddVerificationInfoToResult "PASS" , "Set/Reset PIN is Displayed successfully as expected"
					Else
						iModuleFail = 1
						AddVerificationInfoToResult "FAIL" , "Set/Reset PIN is Not  Displayed successfully"
						Exit Function
					End If

					If  SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebButton("Set/Reset Word and Hint").isenabled =True Then
						CaptureSnapshot()
						AddVerificationInfoToResult "PASS" , "Set/Reset Word and Hint is Displayed successfully as expected"
					Else
						iModuleFail = 1
						AddVerificationInfoToResult "FAIL" , "Set/Reset Word and Hint is Not  Displayed successfully"
						Exit Function
					End If

	End If

	If sDPAValidationInLineItems="DPA Required" OR sDPAValidationInLineItems="Failed" OR sDPAValidationInLineItems="Not Required"Then
			SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms").SiebList("List").SiebPicklist("DPA validation").SiebSelect sDPAValidationInLineItems
			CaptureSnapshot()
				If Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("Set/Reset Password").Exist Then
					CaptureSnapshot()
						AddVerificationInfoToResult "PASS" , "Set/Reset Password is Disabled as expected"
				Else
						iModuleFail = 1
						AddVerificationInfoToResult "FAIL" , "Set/Reset Password is Enabled"
						CaptureSnapshot()
						Exit Function
				End If

				If Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("Set/Reset PIN").Exist Then
							CaptureSnapshot()
						AddVerificationInfoToResult "PASS" , "Set/Reset PIN is Disabled as expected"
						Else
						iModuleFail = 1	
						AddVerificationInfoToResult "FAIL" , "Set/Reset PIN is Enabled"
						CaptureSnapshot()
						Exit Function
				End If

				If Browser("Siebel Call Center_2").Page("Siebel Call Center_2").Frame("View Frame").WebElement("Set/Reset Word and Hint").Exist Then
							CaptureSnapshot()
						AddVerificationInfoToResult "PASS" , "Set/Reset Word and Hint Disabled as expected"
						Else
						iModuleFail = 1
						AddVerificationInfoToResult "FAIL" , "Set/Reset Word and Hint is Enabled"
						CaptureSnapshot()
						Exit Function
				End If

	End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : AddVerifyContacts_fn()
' 	Description : Add and verifies multiple contacts along with adding and verifying account permissions
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################

Public Function AddVerifyContacts_fn()

   Dim adoData	  
	strSQL = "SELECT * FROM [AddVerifyContacts$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sAddContact = adoData( "AddContact")  & ""
	sNewPermission = adoData( "NewPermission")  & ""
	sContact = adoData( "Contact")  & ""
	sPopUp = adoData( "PopUp")  & ""
	sPopUp1 = adoData( "PopUp1")  & ""

	If Ucase(sPopUp)="FALSE" or sPopUp=""Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If

	If Ucase(sPopUp1)="FALSE" or sPopUp1=""Then
		sPopUp1="FALSE"
		sPopUp1=Cbool(sPopUp1)
	End If

		If Not (SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Add").Exist(1)) Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
			Browser("index:=0").Page("index:=0").Sync
		End If
	
	CaptureSnapshot()

	If sAddContact = "Y" Then
	
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Add").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").SiebText("Last name").SiebSetText sContact
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").SiebText("Last name").ProcessKey "EnterKey"
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebButton("OK").SiebClick sPopUp1
		Browser("index:=0").Page("index:=0").Sync
	End If

	If sNewPermission = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebButton("New").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
	End If

	If  sNewPermission="ValidateAccountPermissions" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").ActivateRow 0
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").SiebText("Username").SetText

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebMenu("Menu").Select "Execute Query"
		CaptureSnapshot()
		Exit function
	End If


Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions")
'			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

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

						 If sValue = "Level 0" Then 
'						 SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebMenu("Menu").Select "Execute Query"
							SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Add").SiebClick sPopUp
'								If Browser("Siebel Call Center").Dialog("Siebel").Static("Level 0").Exist(2) Then
'									CaptureSnapshot()
'									AddVerificationInfoToResult  "Info" , "Expected pop occurerred after selecting Level 0"
'									Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
'									Else
'									AddVerificationInfoToResult  "Info" , "Expected pop did not occur  after selecting Level 0"
'							 End if
						End If
					CaptureSnapshot()
				End If
			adoData.MoveNext
			Loop ''
        

CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : VerifyCustComContacts_fn()
' 	Description : This function verifies different fields in Contacts section under Customer Comms
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################

Public Function VerifyCustComContacts_fn()

   Dim adoData	  
	strSQL = "SELECT * FROM [AddVerifyContacts$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Contacts")
			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0


			Do while Not adoData.Eof
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""
				sPopUp = adoData( "PopUp")  & ""
				
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
'				If sUIName <> "" Then
'					 UpdateSiebList objApplet,sUIName,sValue		
'					CaptureSnapshot()
'				End If
			adoData.MoveNext
			Loop ''
        
		CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
		
End Function



'#################################################################################################
' 	Function Name : CustCommsMsgVerify_fn()
' 	Description : This function verifies different fields in Customer Comms section under Account Summary
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################

Public Function CustCommsMsgVerify_fn()

   Dim adoData	  
	strSQL = "SELECT * FROM [AddVerifyContacts$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

			Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Customer Comms List")
			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0


			Do while Not adoData.Eof
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""
				sLocateColValue = adoData( "LocateColValue")  & ""
				sLocateCol = adoData( "LocateCol")  & ""
				sPopUp = adoData( "PopUp")  & ""
				
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
'				If sUIName <> "" Then
'					 UpdateSiebList objApplet,sUIName,sValue		
'					CaptureSnapshot()
'				End If
			adoData.MoveNext
			Loop ''
        
		CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
		
End Function

'''#################################################################################################
'               Function Name : VerifyCustomerSummary_fn
'               Description : 
'   Created By :  Pushkar
'               Creation Date :        
'##################################################################################################
Public Function VerifyCustomerSummary_fn()

                'Get Data
'               Dim adoData        
    strSQL = "SELECT * FROM [VerifyCustomerSummary$] WHERE  [RowNo]=" & iRowNo
                Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'               'sOR
                call SetObjRepository ("Account",sProductDir & "Resources\")

    sOwnedProdService =   adoData("OwnedProdService") & ""
                sUsedProdService =   adoData("UsedProdService") & ""
                sBilledProdService =   adoData("BilledProdService") & ""
                sGoToCustomerSummary =  adoData("GoToCustomerSummary") & ""
                sHistoryCheck =   adoData("HistoryCheck") & ""
'				sResetColumns = adoData("ResetColumns") & ""


                'Flow

                If sGoToCustomerSummary = "Y" Then
                                SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Owned Product and Service view","L3"
                                Browser("index:=0").Page("index:=0").Sync
                End If

    
                'Owned Product and service
                If sOwnedProdService = "Y" Then
                
                                                SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Owned Product and Service view","L3"
                                                Browser("index:=0").Page("index:=0").Sync

                                                Set objApplet =SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience")
                               
                                                sRowCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience").SiebList("List").RowsCount
                                                If sRowCnt > 0   Then 
                                                                     CaptureSnapshot()
																	AddVerificationInfoToResult "Pass","Owned Product and Service link is present"
                                                           ELSE  
                                                                     AddVerificationInfoToResult "Fail","Page is not displayed"
                                                                iModuleFail = 1
                                                End If
                End If

'Used Product and service
                If sUsedProdService = "Y" Then
                
                                               SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Used Product and Service view","L4"
                                                Browser("index:=0").Page("index:=0").Sync
' SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience_2").SiebList("List").DrillDownColumn

                                                                Set objApplet =SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience_2")
'																Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Installed Assets").SiebApplet("Customer Experience_2")

                                                                sRowCnt2= SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience_2").SiebList("List").RowsCount
                                                                If sRowCnt2 >0  Then
                                                                CaptureSnapshot()
																AddVerificationInfoToResult "Pass","Used Product and Service link is present"
                                                                ELSE  
                                                                   AddVerificationInfoToResult "Fail","Page is not displayed"
                                                                   iModuleFail = 1
                                                                End If
                End If

'Billed Product and Service 
                If sBilledProdService = "Y" Then
                                
'    SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience_3").SiebList("List").DrillDownColumn
                                           SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Customer Summary Billed Product and Service view","L4"
                                                                Browser("index:=0").Page("index:=0").Sync

                                                                '
															Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience_3")

                                                                sRowCnt3=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience_3").SiebList("List").RowsCount
                                                If sRowCnt3 > 0  Then
														CaptureSnapshot()
														AddVerificationInfoToResult "Pass","Billed Product and Service link is present"
                                                ELSE  
                                                            AddVerificationInfoToResult "Fail","Page is not displayed"
                                                                iModuleFail = 1
                                End If

                End If

					If ( sResetColumns = "Y" ) Then
						SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Customer Experience_3").SiebMenu("Menu").Select "ColumnsDisplayed"
						If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow_3").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Exist then
									 CaptureSnapshot()
									 Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow_3").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Click
									 Browser("index:=0").Page("index:=0").Sync
                                     AddVerificationInfoToResult "Pass","Reset Defaults Button is clicked successfully"
                        Else
                                     AddVerificationInfoToResult "Fail","Reset Defaults Button is not Enabled"
                                     ModuleFail = 1
					End If
			   End If

                
                                If sHistoryCheck = "Y" Then
                                                Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("History")
                                End If

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

                                CaptureSnapshot()

                If Err.Number <> 0 Then
            iModuleFail = 1
                                                AddVerificationInfoToResult  "Error" , Err.Description
                End If

                End Function


'#################################################################################################
' 	Function Name : CustomerComms_DPAValidation_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CustomerComms_DPAValidation_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AdminData$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\ServerManagement.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")
	
	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"

	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Activities Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
	End If

	Browser("index:=0").Page("index:=0").Sync

	 		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms")

			objApplet.SiebList("List").ActivateRow 0

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
			Loop 

	If Browser("Siebel Call Center").Dialog("Siebel").Static("SIMOFLEX").Exist(5) Then
		Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
	End If

	If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").Exist(5) Then
		sCount = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("DPA validation").GetRoProperty("count")
		For loopVar = 0 to sCount-1
			sItem = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("DPA validation").GetItemByIndex (loopVar)
			If sItem = "DPA Required" OR  sItem = "Failed" OR  sItem = "Not Required" OR  sItem = "Passed" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("Activity Expenses").SiebApplet("Customer Comms").SiebPicklist("DPA validation").Select sItem
				AddVerificationInfoToResult  "Info" ,"DPA Validation item is : " & sItem
				CaptureSnapshot()
			Else
				AddVerificationInfoToResult  "Error" ,"DPA Validation item is : " & sItem
				iModuleFail = 1
			End If
		Next
	Else
		AddVerificationInfoToResult  "Error" ,"ID column is not drilled down successfully."
		iModuleFail = 1
	End If
	

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'''#################################################################################################
'               Function Name : SelectDDTransactionType_fn
'               Description : 
'   Created By :  Pushkar
'               Creation Date :        
'##################################################################################################
Public Function SelectDDTransactionType_fn()

                'Get Data
'               Dim adoData        
    strSQL = "SELECT * FROM [SelectDDTransactionType$] WHERE  [RowNo]=" & iRowNo
                Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'               'sOR
                call SetObjRepository ("Account",sProductDir & "Resources\")

    sDDTransactionType =   adoData("DDTransactionType") & ""
	sClickOnBillingProf = adoData("ClickOnBillingProf") & ""
	sClickonCustComms=  adoData("ClickonCustComms") & ""
	sGotoProfiles = adoData("GotoProfiles") & ""
	sSadminLogin = adoData("SadminLogin") & ""
	If  (sGotoProfiles = "Y")Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
			Browser("index:=0").Page("index:=0").Sync
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebPicklist("DD Transaction Type").SiebSelect sDDTransactionType
			CaptureSnapshot()
			wait 5
	End If

	If  (sGotoProfiles = "Yes" and sSadminLogin = "Yes")Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Customer Profile View","L3"
			Browser("index:=0").Page("index:=0").Sync
			CaptureSnapshot()
	End If

	If  (sClickOnBillingProf = "Yes")Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L4"		
			Browser("index:=0").Page("index:=0").Sync
			CaptureSnapshot()
	End If

	If  (sClickonCustComms = "Y")Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Activities View","L3"
		Browser("index:=0").Page("index:=0").Sync
		CaptureSnapshot()
		wait 15
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities_2").SiebApplet("Customer Comms").SiebMenu("Menu").Select "Run Query"
		wait 20
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities_2").SiebApplet("Customer Comms").SiebMenu("Menu").Select "Run Query"
	End If

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
    End If

	End Function

'''#################################################################################################
'   Function Name : ServiceRequestSubmitSR_fn
'   Description : 
'   Created By :  Tarun
'    Creation Date :        
'##################################################################################################
Public Function ServiceRequestSubmitSR_fn()

'        sOR
    call SetObjRepository ("Account",sProductDir & "Resources\")

    
	SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("Submit SR").SiebClick False

'	sInitialStatus =SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activity").SiebApplet("Service Requests").SiebPicklist("Status").SiebGetROProperty("activeitem")

	wait 5

	SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebMenu("Menu").Select "ExecuteQuery"

	sFinalStatus = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Status").SiebGetROProperty("activeitem")

	AddVerificationInfoToResult  "Info" , "Agent User is able to submit SR and SR Status is changed to " & sFinalStatus


	CaptureSnapshot()

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
    End If

	End Function

'''#################################################################################################
'   Function Name : ServiceRequest_ActivityPlan_fn
'   Description : 
'   Created By :  Tarun
'    Creation Date :        
'##################################################################################################
Public Function ServiceRequest_ActivityPlan_fn()

     strSQL = "SELECT * FROM [ServiceRequest_Activities$] WHERE  [RowNo]=" & iRowNo
	 Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

'        sOR
    call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebScreenViews("ScreenViews").Goto "FS Template Activities View","L3"

	sNewButtonClick = adoData( "ClickNew")  & ""
	sActivityPlan = adoData( "ActivityPlan")  & ""

	If sNewButtonClick = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activity").SiebApplet("Activity Plans").SiebButton("New").SiebClick False
	End If

	If sActivityPlan = "Yes" Then
		rwCnt = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activity").SiebApplet("Activity Plans").SiebList("List").SiebListRowCount
		If rwCnt = 0 Then
			AddVerificationInfoToResult  "Info" ,"No activities plans."
			Exit Function
		Else
			AddVerificationInfoToResult  "Error" ,"Activity Plans exist."
			iModuleFail = 1
			Exit Function
		End If
	End If

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activity").SiebApplet("Activity Plans")

	Do while Not adoData.Eof
		sLocateCol =  adoData("LocateCol") & ""
		sLocateColValue =  adoData("LocateColValue") & ""
		index = 0
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
		If sLocateCol<>"" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,index)
		
			If res = False Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If


	If sUIName <> "" Then
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activity").SiebApplet("Activity Plans").SiebList("List").ActivateRow 0
			UpdateSiebList objApplet,sUIName,sValue
	End If
		adoData.MoveNext
	Loop

	CaptureSnapshot()

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
    End If

	End Function


'''#################################################################################################
'   Function Name : ServiceRequest_Activities_fn
'   Description : 
'   Created By :  Tarun
'    Creation Date :        
'##################################################################################################
Public Function ServiceRequest_Activities_fn()

     strSQL = "SELECT * FROM [ServiceRequest_Activities$] WHERE  [RowNo]=" & iRowNo
	 Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	 sButtonValidation =  adoData("ButtonValidation") & ""
	 sServiceRequestClick =  adoData("ServiceRequestClick") & ""

'        sOR
    call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebScreenViews("ScreenViews").Goto "Service Request Detail View","L3"

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Activities")


	Do while Not adoData.Eof
		sLocateCol =  adoData("LocateCol") & ""
		sLocateColValue =  adoData("LocateColValue") & ""
		index = 0
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
		If sLocateCol<>"" Then

			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,index)
		
			If res = False Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

	If sUIName <> "" Then
			UpdateSiebList objApplet,sUIName,sValue
	End If
		adoData.MoveNext
	Loop

	If sButtonValidation = "Yes" Then
		If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("innertext:=Submit SR","index:=0","class:=miniBtnUICOff").Exist(5) Then
'			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebMenu("Menu").Select "ExecuteQuery"
			AddVerificationInfoToResult  "Info" , "Agent User is not able to submit SR and Submit SR button is disabled when activities status is in progress for mandatory fields."
		Else
			AddVerificationInfoToResult  "Error" , "Submit SR button is not disabled when activities status is in progress for mandatory fields."
			iModuleFail = 1
		End If
	End If

	CaptureSnapshot()

'	sServiceId = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebText("ID").SiebGetROProperty("text")
'
'		If DictionaryTest_G.Exists("ServiceId") Then
'			DictionaryTest_G.Item("ServiceId")=sServiceId
'		else
'			DictionaryTest_G.add "ServiceId",sServiceId
'		End If
'
'		If sServiceRequestClick = "Yes" Then
'			SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Service Request Screen"
'		End If
'	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoView "Personal Service Request List View"


	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
    End If

	End Function

'#################################################################################################
' 	Function Name : BillsEmail_fn(()
' 	Description : This function is used to verify required fields on Customer comms list screen
'   Created By :  Zeba
'	Creation Date :        
'##################################################################################################

Public Function BillsEmail_fn()

   Dim adoData	  
   strSQL = "SELECT * FROM [BillsEmail$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'	sAccountNumber = adoData( "Account_Number")  & ""
''DictionaryTest_G.Item("AccountNo") ="7000215515"
    sAccountNumber = DictionaryTest_G.Item("AccountNo")
'	sAccountName = adoData( "AccountName")  & ""

	sPopUp = adoData( "Popup")  & ""
	sAction = adoData( "Action")  & ""
	sEmail = adoData("Email")
	sEmail1 = adoData("Email1")
	sEmail2 = adoData("Email2")
'	sCapture = adoData( "Capture")  & ""
'	sUIName = adoData( "UIName")  & ""


	If Ucase(sPopUp)="FALSE" or sPopUp=""Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If

'	If sAccountNumber = "" Then
'		AddVerificationInfoToResult "Fail","Account number is null and not retrieved from data base"
'		iModuleFail = 1
'		Exit Function
'	End If


	If (sAction="WithoutEmail") Then

    	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebDrillDownColumn "Name",0
		Browser("index:=0").Page("index:=0").Sync

		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoView "CMU Billing Invoice View"
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Bills").SiebButton("Email Copy Bill").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Select Email Address").SiebButton("Send Email").SiebClick sPopUp
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Select Email Address").SiebButton("Cancel").SiebClick False
		AddVerificationInfoToResult "PASS" , "Popup came as expected"

	End If


 	If (sAction="UpdateEmail") Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebButton("Update Email").Click
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Siebel").SiebText("Email").SiebSetText sEmail1
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Siebel").SiebPicklist("SiebPicklist").SetText sEmail2
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Siebel").SiebText("Re-type Email").SiebSetText sEmail1
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Siebel").SiebPicklist("SiebPicklist_2").SetText sEmail2
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Siebel").SiebButton("OK").SiebClick False

		 SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").SiebDrillDownColumn "Name",0
		Browser("index:=0").Page("index:=0").Sync

		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoView "CMU Billing Invoice View"

        	CaptureSnapshot()
			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Bills").SiebButton("Email Copy Bill").SiebClick False
			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Select Email Address").SiebPicklist("Select Email address").SiebSelect "Account holder Email"
			CaptureSnapshot()
			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Select Email Address").SiebButton("Send Email").SiebClick sPopUp
'			If Browser("Siebel Call Center_2").Dialog("Siebel").Static("Your request has been").Exist Then
'				Browser("Siebel Call Center_2").Dialog("Siebel").Static("Your request has been").Click
'			End If
			Browser("index:=0").Page("index:=0").Sync
			wait 20
			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Billing profile").SiebButton("Customer Account").Click

     

	End If

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

	End Function


'#################################################################################################
' 	Function Name : ServiceRequestScreenMain_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By : Tarun
'	Creation Date :        
'##################################################################################################
Public Function ServiceRequestScreenMain_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [ServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""
	sOwnerValidation = adoData( "OwnerValidation")  & ""
	sChangeOwner = adoData( "ChangeOwner")  & ""
	sOwnerLogin = adoData( "OwnerLogin")  & ""
	sServiceIdValidation = adoData( "ServiceIdValidation")  & ""
	sSourceFieldValidation = adoData( "SourceFieldValidation")  & ""
	sMyServReq = adoData( "MyServReq")  & ""
	sAllOrgServReq = adoData( "AllOrgServReq")  & ""
	sCustomerAccountClick = adoData( "CustomerAccountClick")  & ""

	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")


	If SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GetRoProperty("ActiveScreen") <>  "Service Request Screen" Then
		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Service Request Screen"
	End If


	Browser("index:=0").Page("index:=0").Sync


		If sMyServReq = "Y" Then
			Set ServiceRequestView = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests")
			Set obj = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("All Service Request across Organizations").SiebApplet("Service Requests")
			If ServiceRequestView.SiebApplet("Service Requests").Exist(5) Then
				AddVerificationInfoToResult  "Info" , "Service Request screen is displayed successfully."
			Else
				AddVerificationInfoToResult  "Error" , "Service Request screen is not displayed successfully."
				iModuleFail = 1
				Exit Function
			End If
			ElseIf  sAllOrgServReq = "Y" Then
			Set ServiceRequestView  = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("All Service Request across")
			Set obj = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("All Service Request across").SiebApplet("Service Requests")
				If ServiceRequestView.SiebApplet("Service Requests").Exist(5) Then
					AddVerificationInfoToResult  "Info" , "Service Request screen is displayed successfully."
				Else
					AddVerificationInfoToResult  "Error" , "Service Request screen is not displayed successfully."
					iModuleFail = 1
					Exit Function
				End If
		End If

	
	If sServiceIdValidation = "Y" Then
		ServiceRequestView.SiebApplet("Service Requests").SiebMenu("Menu").Select "NewQuery"
		ServiceRequestView.SiebApplet("Service Requests").SiebList("List").SiebText("ID").SiebSetText DictionaryTest_G.Item("ServiceId")
		ServiceRequestView.SiebApplet("Service Requests").SiebButton("Go").SiebClick False
		CaptureSnapshot()
	End If

	If sOwnerValidation = "Y" Then
		sServiceId = ServiceRequestView.SiebApplet("Service Requests_2").SiebText("ID").GetRoProperty("text")
		DictionaryTest_G.Item("ServiceId") = sServiceId
		sCreatedBy = ServiceRequestView.SiebApplet("Service Requests_2").SiebText("Created by").GetRoProperty("text")
		sOwner = ServiceRequestView.SiebApplet("Service Requests_2").SiebText("Owner").GetRoProperty("text")

		If  sServiceIdValidation = "Y" Then
			rwCnt = ServiceRequestView.SiebApplet("Service Requests").SiebList("List").RowsCount
			If rwCnt = "1" Then
				AddVerificationInfoToResult  "Info" , "Service Request is created successfully when Created by is  " & sCreatedBy & " Owner is  " & sOwner
			Else
				AddVerificationInfoToResult  "Error" , "Service Request is not  created successfully when Created by is  " & sCreatedBy & " Owner is  " & sOwner
				iModuleFail = 1
			End If
		End If

	End If

	If Ucase(sClickNew) = "Y" Then
		ServiceRequestView.SiebApplet("Service Requests").SiebButton("New").SiebClick sPopUp
		If sOwnerValidation = "Y" Then
			If sCreatedBy = sOwner Then
				AddVerificationInfoToResult  "Info" , "Error message is disaplyed successfully when both Created by account - " & sCreatedBy & " and Owner account - " & sOwner & " acccount are same and Service Id created is : " & DictionaryTest_G.Item("ServiceId")
			End If
		End If
	End If

'	ServiceRequestView.SiebApplet("Service Requests").SiebList("List").ActivateRow 0

    
	Do while Not adoData.Eof
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	

		If sValue = "NOW" Then
			TimeRequested = FormatDateTime (Time, vbShortTime)& Right (Time, 6)
			MN =Left( MonthName(Month(Date)),3)
			DateRequested = Date					
			DateRequested = day(DateRequested) & "/" & MN & "/" &  year(DateRequested) &" "& TimeRequested
			sValue = DateRequested								
		End If

		If sUIName <> "" Then
			UpdateSiebList obj,sUIName,sValue
		End If

		If instr (sUIName, "Contact") > 0 Then
			If ServiceRequestView.SiebApplet("Pick Contact").Exist(5) Then
				ServiceRequestView.SiebApplet("Pick Contact").SiebButton("OK").SiebClick False
			Else
				AddVerificationInfoToResult  "Error" , "Accounts Open pop up is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If
		End If
	
		If instr (sUIName, "Account name") > 0 Then
			If ServiceRequestView.SiebApplet("Pick Account").Exist(5) Then
				ServiceRequestView.SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				CaptureSnapshot()
			Else
				AddVerificationInfoToResult  "Error" , "Account name Open pop up is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If
		End If
		adoData.MoveNext
	Loop

	If sChangeOwner = "Y" Then
		CaptureSnapshot()
		ServiceRequestView.SiebApplet("Service Requests_2").SiebText("Owner").OpenPopup
		If ServiceRequestView.SiebApplet("Pick Service Request Owner").Exist(5) Then
			ServiceRequestView.SiebApplet("Pick Service Request Owner").SiebPicklist("SiebPicklist").SiebSelect "User ID"
			ServiceRequestView.SiebApplet("Pick Service Request Owner").SiebText("SiebText").SiebSetText sOwnerLogin
			ServiceRequestView.SiebApplet("Pick Service Request Owner").SiebButton("Go").SiebClick False
		Else
			AddVerificationInfoToResult  "Error" , "Owner Open pop up is not clicked successfully."
			iModuleFail = 1
		End If
	
	End If

	If sSourceFieldValidation = "Y"  Then
		ServiceRequestView.click
	End If

	If sCustomerAccountClick = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("Customer Account").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
	End If

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : Contacts_AccountPermission_fn
' 	Description : This function is used to used to click on contacts tab on accounts page and drilldown on last name.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Contacts_AccountPermission_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AccountPermission$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sOnlineAccountValidation = adoData( "OnlineAccountValidation")  & ""
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sType = adoData( "Type")  & ""
	sAdd = adoData( "Add")  & ""
	sPrimaryCheck = adoData( "PrimaryCheck")  & ""

	'Flow
	activeView = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Account Detail - Contacts View" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
	End If

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").Exist(5) Then
		AddVerificationInfoToResult "Info","Contacts tab clicked successfully. "
	Else
		AddVerificationInfoToResult "Error","Contacts tab is not clicked successfully. "
		iModuleFail = 1
		Exit Function
	End If

	If sAdd = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Add").SiebClick False
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").Exist(5) Then
				AddVerificationInfoToResult  "Info" , "Add Contact button is clicked successfully."
			Else
				AddVerificationInfoToResult  "Error" , "Add click button is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").ActivateRow 0

			Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts")

		Do while Not adoData.Eof
	
			sUIName = adoData( "UIName")  & ""
			sValue = adoData( "Value")  & ""	
	
			If sUIName <> "" Then
				UpdateSiebList objApplet,sUIName,sValue
			End If
			adoData.MoveNext
		Loop

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").SiebText("Last name").ProcessKey "EnterKey"
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").ActivateRow 0
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebButton("OK").SiebClick False
		Exit Function

	End If

	If  sType = "Contacts" Then
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts")
		objApplet.SiebList("List").ActivateRow 0
	ElseIf sType = "Permission" Then
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions")
			rwCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebList("List").RowsCount
			If rwCnt <> 0 Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebList("List").ActivateRow 0
			End If

			If rwCnt = 0 AND sPrimaryCheck ="No" Then
				AddVerificationInfoToResult  "Info" ,"Value captured from account permission page is blank when primary check box is unchecked for new contact."
'				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebList("List").highlight
				CaptureSnapshot()
				Exit Function
			End If

			If rwCnt = 0 AND sPrimaryCheck ="Yes" Then
				AddVerificationInfoToResult  "Info" ,"Value captured from account permission page is blank when primary check box is checked for new contact."
'				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebList("List").highlight
				CaptureSnapshot()
				Exit Function
			End If
	End If

	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	

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
	End If

	If instr(sUIName,"CaptureTextValue") > 0 And Instr(sValue,"A person who has been granted the") > 0 Then
		AddVerificationInfoToResult  "Info" ,"Value captured from account permission page - " & sValue & " for main user."
	End If
		adoData.MoveNext
	Loop

		CaptureSnapshot()			
  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : Addresses_ColumnsDisplayed_fn
' 	Description :  This function is used to check whether Delete button is enabled or disabled at addresses tab
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Addresses_ColumnsDisplayed_fn()

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Business Address View","L3"

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").SiebMenu("Menu").Select "ColumnsDisplayed"
		If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").Exist(5) Then
			sItems =Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebList("ShownItems").GetROProperty("all items")
			If Instr(sItems,"For the attention of") > 0 Then
				AddVerificationInfoToResult  "Error" ,"For the attention of value is present in columns displayed."
				iModuleFail = 1
			Else
				AddVerificationInfoToResult  "Info" ,"For the attention of value is not present in columns displayed."
			End If
		Else
			AddVerificationInfoToResult  "Error" ,"Columns displayed is not selected successfully from menu item."
			iModuleFail = 1
			Exit Function
		End If

	Else
		AddVerificationInfoToResult  "Error" ,"Addresses tab is not clicked successfully."
		iModuleFail =1
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : Contacts_AccountPermission_NewButton_fn
' 	Description : This function is used to used to click on contacts tab on accounts page and drilldown on last name.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Contacts_AccountPermission_NewButton_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AccountPermission$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sOnlineAccountValidation = adoData( "OnlineAccountValidation")  & ""
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sNewButton = adoData( "NewButton")  & ""
	sAdd = adoData( "Add")  & ""
	sType = adoData( "Type")  & ""
	

	'Flow
	activeView = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Account Detail - Contacts View" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
	End If

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").Exist(5) Then
		AddVerificationInfoToResult "Info","Contacts tab clicked successfully. "
	Else
		AddVerificationInfoToResult "Error","Contacts tab is not clicked successfully. "
		iModuleFail = 1
		Exit Function
	End If

	If sNewButton = "Click" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebButton("New").SiebClick False
	End If

	If sAdd = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebButton("Add").SiebClick False
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").Exist(5) Then
				AddVerificationInfoToResult  "Info" , "Add Contact button is clicked successfully."
			Else
				AddVerificationInfoToResult  "Error" , "Add click button is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").ActivateRow 0

			Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts")

		Do while Not adoData.Eof
	
			sUIName = adoData( "UIName")  & ""
			sValue = adoData( "Value")  & ""	
	
			If sUIName <> "" Then
				UpdateSiebList objApplet,sUIName,sValue
			End If
			adoData.MoveNext
		Loop

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").SiebText("Last name").ProcessKey "EnterKey"
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebList("List").ActivateRow 0
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Add Contacts").SiebButton("OK").SiebClick False
		Exit Function

	End If

	If  sType = "Contacts" Then
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts")
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts").SiebList("List").ActivateRow 0
	ElseIf sType = "Permission" Then
		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions")
'		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebList("List").ActivateRow 0
	End If

	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""	

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
	End If

		adoData.MoveNext
	Loop

	If sNewButton = "Disabled" Then
		If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("New").Exist(5) Then
			AddVerificationInfoToResult  "Info" , "New button is disabled for Primary contact."
		Else
			AddVerificationInfoToResult  "Error" , "New button is enabled for Primary contact. It should be disabled"
			iModuleFail = 1
		End if
	End If

	If sNewButton = "Enabled" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Permissions").SiebButton("New").Exist(5) Then
			CaptureSnapshot()
			AddVerificationInfoToResult  "Info" , "New button is enabled for Secondary contact when permission is not assigned to Secondary User."
		Else
			AddVerificationInfoToResult  "Error" , "New button is disabled for Secondary contact when permission is not assigned to Secondary User.. It should be enabled."
			iModuleFail = 1
		End if
	End If

	If sNewButton = "ClickDisabled" Then
		If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("New").Exist(5) Then
			AddVerificationInfoToResult  "Info" , "New button is disabled for Secondary contact when permission is granted to user."
		Else
			AddVerificationInfoToResult  "Error" , "New button is enabled for Secondary contact when permission is granted to user.. It should be disabled"
			iModuleFail = 1
		End if
	End If


		CaptureSnapshot()			
  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : ValidateCustComms_fn()
' 	Description : This function is used to verify required fields on Customer comms list screen
'   Created By :  
'	Creation Date :        
'##################################################################################################

Public Function ValidateCustComms_fn()

   Dim adoData	  
	strSQL = "SELECT * FROM [ValidateCustComms$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
'	
'	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

'		sLocateCol = adoData( "LocateCol")  & ""
		sViews = adoData( "Views")  & ""
		sMain = adoData( "Main")  & ""

		
'			SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
'			Browser("index:=0").Page("index:=0").Sync

			If sViews = "Customer Comms" Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Activities View","L3"
				Browser("index:=0").Page("index:=0").Sync
				CaptureSnapshot()
				Set objApplet  = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Activities_2").SiebApplet("Customer Comms")
			End If

		If sMain = "Y"  Then
			SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"
			Browser("index:=0").Page("index:=0").Sync
			If sViews = "My Customer Comms" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Activity List View","L2"
				Browser("index:=0").Page("index:=0").Sync
				CaptureSnapshot()
				Set objApplet  = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms")
			ElseIf sViews = "My Team's Customer Comms" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Manager's Activity List View","L2"
'				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "Manager's Activity List View","L2"
				Browser("index:=0").Page("index:=0").Sync
				CaptureSnapshot()
				Set objApplet  = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Customer Comms")
			ElseIf 	sViews = "All Customer Comms across My Organization" Then
				SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebScreenViews("ScreenViews").Goto "VF All Activity List View Across My Organization","L2"
				Browser("index:=0").Page("index:=0").Sync
				CaptureSnapshot()
				Set objApplet  = SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities_2").SiebApplet("Activities")
			End If
		  End If


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
								If  Cstr(res)="True" Then
									AddVerificationInfoToResult  "Info" , sLocateCol &  " found in the " & sViews
								Else
									iModuleFail = 1
									AddVerificationInfoToResult  "Error" , sLocateCol & " not found in the " & sViews
									Exit Function
								End If
						End If    
						If sUIName <> "" Then
							UpdateSiebList objApplet,sUIName,sValue
						End If
					adoData.MoveNext
					Loop ''

		
                            
		If Err.Number <> 0 Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , Err.Description
		End If
						
	End Function
 
'#################################################################################################
' 	Function Name : SRResolved_CustomerAccount_fn
' 	Description :  This function is used to click on new button in Service accounts applet on account s page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function SRResolved_CustomerAccount_fn()


	Dim adoData	  
	strSQL = "SELECT * FROM [CreateNewServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sButton = adoData( "Button")  & ""

	sPopUp = adoData( "Popup")  & ""

	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If

	'Flow

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").Exist(5) Then
		AddVerificationInfoToResult  "Info" , "SR Resolved page is displayed successfully."
	Else
		AddVerificationInfoToResult  "Error" , "IDis not drilled down successfully and SR Resolved page is not displayed successfully."
		iModuleFail = 1
		Exit Function
	End If

    If sButton = "Resolved" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("SR Resolved").SiebClick sPopUp
		CaptureSnapshot()
	End If

	If sButton = "CustomerAccount" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("Customer Account").SiebClick sPopUp
	End If

    If sButton = "Validation" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("SR Resolved").SiebClick False
		wait 2
		sStatus = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Status").GetRoProperty("activeitem")
		If sStatus = "Closed" Then
			AddVerificationInfoToResult  "Info" , "User is able to click on SR Resolved button when logined from another user and SR status is closed."
		Else
			AddVerificationInfoToResult  "Error" , "User is not able to click on SR Resolved button when logined from another user."
			iModuleFail = 1
		End If
		CaptureSnapshot()
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : ServiceRequest_Activity_fn
' 	Description : This function is used to Enter Installed id and Verify Specific Columns
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ServiceRequest_Activity_fn()

	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

		If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Activities").Exist(5) Then
			AddVerificationInfoToResult  "Info" ,"Service Request Activities page is displayed successfully."

			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Activities").SiebMenu("Menu").Select "ColumnsDisplayed"
			If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Exist(5) Then
				Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Reset Defaults").Click
				Browser("index:=0").Page("index:=0").Sync
			Else
				AddVerificationInfoToResult  "Error" ,"Columns displayed is not selected successfully from menu items."
				iModuleFail = 1
				Exit Function
			End If
		Else
			AddVerificationInfoToResult  "Error" ,"Service Request Activities page is not displayed successfully."
			iModuleFail = 1
			Exit Function
		End If


		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Activities")
		colCnt = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Activities").SiebList("List").ColumnsCount

					For loopVar = 0 to colCnt
						sReposName = objApplet.SiebList("List").GetColumnRepositoryNameByIndex(loopVar)
						sUIName = objApplet.SiebList("List").GetColumnUIName(sReposName)
						If sUIName = "Type" Then
							val1 = loopVar
						End If
				
						If sUIName = "Alarm" Then
							val2 = loopVar
						End If
				
						If sUIName = "Description" Then
							val3 = loopVar
							Exit For
						End If
					
					Next
'				
				
					If Cint(val2) = Cint(val1) + 1 AND Cint(val2) = Cint(val3) - 1 Then
						AddVerificationInfoToResult  "Info" ,"Alarm column is in between Type and Description column."
					Else
						AddVerificationInfoToResult  "Error" ,"Alarm column is not in between Type and Description column."
						iModuleFail = 1
					End If

					CaptureSnapshot()
    
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : ServiceRequestInformation_fn
' 	Description : This function is used to click new button on service request and perform the actions 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ServiceRequestInformation_fn()

	'Get Data
	Dim adoData	  
	strSQL = "SELECT * FROM [CreateNewServiceRequest$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	sClickNew = adoData( "ClickNew")  & ""

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").Exist(5) Then
		If sClickNew = "Disabled" Then
			sEditable = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").SiebPicklist("Request Source").GetRoProperty("isenabled")
			If sEditable = "False" Then
				AddVerificationInfoToResult  "Info" , "Request Source picklist is disabled when no value is selected in Type list."
			Else
				AddVerificationInfoToResult  "Error" , "Request Source picklist is enabled when no value is selected in Type list. It shpuld be disabled."
				iModuleFail = 1
			End If
		End If

		If sClickNew = "Enabled" Then
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").SiebPicklist("Type").Select "Customer Relations"
			sEditable = SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("Service Requests_2").SiebPicklist("Request Source").GetRoProperty("isenabled")
			If sEditable = "True" Then
				AddVerificationInfoToResult  "Info" , "Request Source picklist is enabled after selecting  Customer Relations from Type list."
			Else
				AddVerificationInfoToResult  "Error" ,  "Request Source picklist is diasbled after selecting  Customer Relations from Type list. It should be enabled."
				iModuleFail = 1
			End If
		End If
	Else
		AddVerificationInfoToResult  "Error" , "Service Request Information page is not displayed successfully."
		iModuleFail = 1
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function