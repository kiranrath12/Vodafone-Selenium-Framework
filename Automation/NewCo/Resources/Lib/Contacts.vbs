'''''#################################################################################################
' 	Function Name : RetrieveAccountDetails_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function RetrieveDetailsAccount_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [CreateNewAccount$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

		Set AccCont = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("First name")
		Set AccSumm =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("First name")

		
		If AccCont.Exist Then
			sFirst_Name = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("First name").GetROProperty("text")
			sLast_Name =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Account Details").SiebText("Last name").GetROProperty("text")
		Else
			sFirst_Name = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("First name").GetROProperty("text")
			sLast_Name = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Details").SiebText("Last name").GetROProperty("text")
		End If




			If DictionaryTest_G.Exists("FirstName") Then
				DictionaryTest_G.Item("FirstName") = sFirst_Name
			else
				DictionaryTest_G.add "FirstName",sFirst_Name
			End If
	
			If DictionaryTest_G.Exists("LastName") Then
				DictionaryTest_G.Item("LastName") = sLast_Name
			else
				DictionaryTest_G.add "LastName",sLast_Name
			End If
CaptureSnapshot()
	End Function

'#################################################################################################
' 	Function Name : OpenContactList_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function OpenContactList_fn()

   call SetObjRepository ("Contacts",sProductDir & "Resources\")


   		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Contacts View","L3"
		Browser("index:=0").Page("index:=0").Sync

		CaptureSnapshot()

	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Contacts").SiebApplet("Contacts")
'		UpdateSiebList obj,sUIName,sValue
		sUIName = "DrillDown|Last name"
		sValue = DictionaryTest_G.Item("LastName")
		UpdateSiebList objApplet,sUIName,sValue
		Browser("index:=0").Page("index:=0").Sync
		CaptureSnapshot()
		
End Function

'#################################################################################################
' 	Function Name : AuditTrailContacts_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function AuditTrailContacts_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [AuditTrailContact$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Contacts.xls",strSQL)
		
			sAccountAudit = adoData("AccountAudit") & ""

	'sOR
	call SetObjRepository ("Contacts",sProductDir & "Resources\")

		If sAccountAudit = "Y" Then

			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF Account - Audit Trail View","L3"
			Browser("index:=0").Page("index:=0").Sync
			CaptureSnapshot()
			Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Assets").SiebApplet("Read Audit Items")

		Else

			SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebScreenViews("ScreenViews").Goto "VF Contact - Audit Trail View","L3"
			Browser("index:=0").Page("index:=0").Sync
			CaptureSnapshot()
			Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Contacts").SiebView("Account Assets").SiebApplet("Read Audit Items")
			objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList").ActivateRow 0

		End If


			Do while Not adoData.Eof
				sUIName = adoData( "UIName")  & ""
				sValue = adoData( "Value")  & ""
					If sValue = "Today*" Then
						MN =Left( MonthName(Month(Date)),3)
						DateRequested = Date					
						DateRequested = day(DateRequested) & "/" & MN & "/" &  year(DateRequested)
						sValue = DateRequested&"*"
						sLocateColValue = sValue
					Else
					sLocateColValue = adoData( "LocateColValue")  & ""
					End If
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
End Function


