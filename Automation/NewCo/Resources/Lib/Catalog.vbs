
'#################################################################################################
' 	Function Name : CatalogueSearch_fn
' 	Description : This function searches product or promotion in catalogue search
'   Created By :  Automation
'	Creation Date :        
'##################################################################################################
Public Function CatalogueSearch_fn()

'Get Data
	Dim adoData	  
'    strSQL = "SELECT * FROM [CatalogueSearch$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)


	
	'sOR
	call SetObjRepository ("Catalog",sProductDir & "Resources\")

	'Flow


'	sProductName = adoData(sEnv) & ""

	sProductName = DictionaryTest_G.Item("PartNo")

	Browser("index:=0").Page("index:=0").Sync
	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview") <> "Sales Order-Browse Catalog View" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Sales Order-Browse Catalog View","L3"
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Sales Order-Search Catalog View","L4"
	
	Browser("index:=0").Page("index:=0").Sync
	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebEdit("Product_20_ID").WebSet sProductName

	Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Search").WebClick

	Browser("index:=0").Page("index:=0").Sync
	row =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Search").SiebApplet("Results").SiebList("List").SiebListRowCount
	
	If row > 0 Then
		AddVerificationInfoToResult  "Product Search" , "Search passed"
	else
		AddVerificationInfoToResult  "Error" , "Product not present"
		 iModuleFail = 1
		 Exit Function
	End If

	
	
	
'    SiebApplication("Siebel Call Center").SetTimeOut(2000)
'    Reporter.Filter = 3
'	On error resume next
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Search").SiebApplet("Results").SiebButton("Add Item").SiebClick False
'    Reporter.Filter = 0
'    Err.Clear

	'Browser("index:=0").Page("index:=0").Sync
	
	row1 = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Search").SiebApplet("Line Items").SiebList("List").SiebListRowCount
	
	If row1 > 0 Then
		AddVerificationInfoToResult  "Pass" , "Product Added successfully"
'	else
'		AddVerificationInfoToResult  "Error" , "Product not Added successfully"
'		 iModuleFail = 1
	End If



	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name : OrderSummary_fn
' 	Description : This function compares the value of One oFF cost  and LLD
'   Created By :  Tarun
'	Creation Date :        20/02/2015
'##################################################################################################
Public Function OrderSummary_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [OrderSummary$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	
	'sOR
	call SetObjRepository ("Catalog",sProductDir & "Resources\")

	'Flow

	strLLD =  adoData("LLD") & ""

	strOneOffCost = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Search").SiebApplet("OrdersSummary").SiebCurrency("One-off cost").SiebGetRoProperty ("text")
	
	If Trim(strOneOffCost) =  strLLD Then
		AddVerificationInfoToResult  "One Off Cost & LLD- PASS" , "Both the values are matching."
	Else
		AddVerificationInfoToResult  "One Off Cost & LLD- FAIL" , "Both the values are not matching."
		iModuleFail = 1
	End If
	
	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : FastOrdersProducts_fn
' 	Description : This function compares the value of One oFF cost  and LLD
'   Created By :  Ankit
'	Creation Date :        16/06/2016
'##################################################################################################
Public Function FastOrdersProducts_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [FastOrdersProducts$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	
	'sOR
	call SetObjRepository ("Catalog",sProductDir & "Resources\")

	'Flow

	sProductType =  adoData("ProductType") & ""
'	sLocateCol =  adoData("LocateCol") & ""
'	sLocateColValue =  adoData("LocateColValue") & ""
'	index = 0

'	Set oDict = CreateObject("Scripting.Dictionary")
'
'	If instr(sLocateColValue,"Promotion")>0 Then
'		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
'	End If
'	arr1 = Split (sLocateCol,"|")
'	arr2 = Split (sLocateColValue,"|")
'	For i=0 to Ubound (arr1)
'		oDict.Add arr1(i),arr2(i)
'
'	Next

	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Add Remove Products").SiebPicklist("Product Type").SiebSelect sProductType
	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Add Remove Products").SiebButton("Go").Click
'	CaptureSnapshot()
	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Add Remove Products")
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
'	Set objSiebList = SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Add Remove Products").SiebList("micclass:=SiebList","repositoryname:=SiebList")


		UpdateSiebList objApplet,sUIName,sValue
		CaptureSnapshot()
		adoData.MoveNext
	Loop
	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Add Remove Products").SiebButton("OK").Click

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : BillingProfileNameSelection_fn
' 	Description : This function is used to click on Profile tab on Account page and click on Name Billing Profile from SiebLIst
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BillingProfileNameSelection_fn()

	'Get Data
'	Dim adoData	  
    strSQL = "SELECT * FROM [BillingProfileNameSelection$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3" ' clicking on Profile tab

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L4"  ''click on billing profile tab


		Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile")


	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
		End If
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
		sCheckPaymentDate = adoData("CheckPaymentDate")  & ""
		sPopUp = adoData("PopUp")  & ""

			If sCheckPaymentDate = "Y" Then
					 SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebButton("Check Payment Date").SiebClick sPopUp
					  Browser("index:=0").Page("index:=0").Sync
				      CaptureSnapshot()
			End If

		If sLocateCol <> "" Then
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" ,sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

		If sUIName <> "" Then		
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


'#################################################################################################
' 	Function Name : OrderBillingServiceAccountSelection_fn
' 	Description :  This function is used to select billing account and service account from orders page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderBillingServiceAccountSelection_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [BillingServiceAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow
	

	sAccountSelection = adoData( "AccountSelection")  & ""
	sAccountSet = adoData( "AccountSet")  & ""
	sFind = adoData( "Find")  & ""

	If  sAccountSet = "ACCOUNT NO" Then
		sAccountSet = DictionaryTest_G("AccountNo")
	ElseIf sAccountSet = "CHILDACCOUNT" Then
		sAccountSet = DictionaryTest_G.Item("PrePostAccountNo")
	End If

	Browser("index:=0").Page("index:=0").Sync

		If sAccountSelection = "Billing Account" Then
			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").OpenPopup
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebPicklist("SiebPicklist").SiebSelect sFind
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebText("SiebText").SiebSetText sAccountSet
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
				CaptureSnapshot()
				If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Exist(3) Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				End If
				
				AddVerificationInfoToResult  "Info" , "Billing Account is selected successfully."
			Else
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").click
				Browser("index:=0").Page("index:=0").Sync
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").OpenPopup
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebPicklist("SiebPicklist").SiebSelect sFind
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebText("SiebText").SiebSetText  sAccountSet
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
				CaptureSnapshot()
				If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Exist(3) Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				End If
				AddVerificationInfoToResult  "Info" , "Billing Account is selected successfully."
			End If
		ElseIf sAccountSelection = "Service Account" Then
			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").OpenPopup
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebPicklist("SiebPicklist").SiebSelect sFind
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebText("SiebText").SiebSetText sAccountSet
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
				CaptureSnapshot()
				If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Exist(3) Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				End If
				AddVerificationInfoToResult  "Info" , "Service Account is selected successfully."
			Else
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").click
				Browser("index:=0").Page("index:=0").Sync
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").OpenPopup
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebPicklist("SiebPicklist").SiebSelect sFind
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebText("SiebText").SiebSetText  sAccountSet
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
				CaptureSnapshot()
				If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Exist(3) Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				End If
				AddVerificationInfoToResult  "Info" , "Service Account is selected successfully."
			End If
		ElseIf sAccountSelection = "Service AccountOrderLineItem" Then
			If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebPicklist("SiebPicklist").SiebSelect sFind
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebText("SiebText").SiebSetText sAccountSet
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Go").SiebClick False
				CaptureSnapshot()
				If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Exist(3) Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").SiebClick False
				End If
				AddVerificationInfoToResult  "Info" , "Service Account is selected successfully."	
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
' 	Function Name : OrderSubAccountsBillingProfileSelection_fn
' 	Description :  This function is used to select billing profile while doing Pre/Post migration in Orders page
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderSubAccountsBillingProfileSelection_fn()


'	Dim adoData	  
    strSQL = "SELECT * FROM [BillingAccountEnableDisable$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

    sCase = adoData( "Case")  & ""

	Browser("index:=0").Page("index:=0").Sync



	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").OpenPopup

		Select Case sCase

		Case "Disabled"
	
			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebList("List").Exist(5) Then
				If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebElement("innertext:=OK","index:=0","class:=miniBtnUICOff").Exist(3) Then
					AddVerificationInfoToResult  "Info" , "Ok Button is disabled for Billing Account"
					CaptureSnapshot()
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Cancel").SiebClick False
				Else
					AddVerificationInfoToResult  "Fail" , "Ok Button is enabled for Billing Account."
					iModuleFail = 1
				End If
			Else
				AddVerificationInfoToResult  "Fail" , "Billing Account pop up button is not clicked sucessully."
				iModuleFail = 1
				Exit Function
			End If

		Case "Enabled"

			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebList("List").Exist(5) Then
				If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Exist(3) Then
					AddVerificationInfoToResult  "Info" , "Ok Button is Enabled for Billing Account."
					CaptureSnapshot()
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Cancel").SiebClick False
				Else
					AddVerificationInfoToResult  "Fail" , "Ok Button is disabled for Billing Account"
					iModuleFail = 1
				End If
			Else
				AddVerificationInfoToResult  "Fail" , "Billing Account pop up button is not clicked sucessully."
				iModuleFail = 1
				Exit Function
			End If
	
		End Select

	Else
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").click
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").OpenPopup

		Select Case sCase

		Case "Disabled"
	
			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebList("List").Exist(5) Then
				If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebElement("innertext:=OK","index:=0","class:=miniBtnUICOff").Exist(3) Then
					AddVerificationInfoToResult  "Info" , "Ok Button is disabled for Billing Account"
					CaptureSnapshot()
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Cancel").SiebClick False
				Else
					AddVerificationInfoToResult  "Fail" , "Ok Button is enabled for Billing Account."
					iModuleFail = 1
				End If
			Else
				AddVerificationInfoToResult  "Fail" , "Billing Account pop up button is not clicked sucessully."
				iModuleFail = 1
				Exit Function
			End If

		Case "Enabled"

			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebList("List").Exist(5) Then
				If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Exist(3) Then
					AddVerificationInfoToResult  "Info" , "Ok Button is Enabled for Billing Account."
					CaptureSnapshot()
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("Cancel").SiebClick False
				Else
					AddVerificationInfoToResult  "Fail" , "Ok Button is disabled for Billing Account"
					iModuleFail = 1
				End If
			Else
				AddVerificationInfoToResult  "Fail" , "Billing Account pop up button is not clicked sucessully."
				iModuleFail = 1
				Exit Function
			End If
	
		End Select

	End If


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : SubAccountBillingAccountNameSelection_fn
' 	Description :  This function is used to click on Bulk modify and perform desired action
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function SubAccountBillingAccountNameSelection_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [BillingAccountNameSelection$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Browser("index:=0").Page("index:=0").Sync

    Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Sub Accounts").SiebApplet("Billing Accounts")


	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
		End If
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""

		If sLocateCol <> "" Then
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" ,sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

		If sUIName <> "" Then		
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

'#################################################################################################
' 	Function Name : AgreementsAccountNameSelection_fn
' 	Description :  This function is used to drilldown on Account name in agreements tab
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AgreementsAccountNameSelection_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [AgreementAccountameSelect$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Browser("index:=0").Page("index:=0").Sync

    Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Agreements").SiebApplet("Agreements List View")


	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
		End If
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""

		If sLocateCol <> "" Then
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" ,sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

		If sUIName <> "" Then		
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

'#################################################################################################
' 	Function Name : AgreementsConditionalCharges_fn
' 	Description :  This function is used to click on conditional charges in Agreements tab
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AgreementsConditionalCharges_fn()


'	Dim adoData	  
'    strSQL = "SELECT * FROM [AgreementAccountameSelect$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebScreenViews("ScreenViews").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebScreenViews("ScreenViews").Goto "FS Agree Conditional Charge View","L4"  ' clicking on conditional charges
	Else
		iModuleFail = 1
		AddVerificationInfoToResult  "Fail" ,"Account name is not drilled down successfully in Agreements tab."
	End If
 
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : AgreementsConditionalChargesMPSSelect_fn
' 	Description :  This function is used to select MPS in conditional charges in Agreements tab and select about  recors from menu items
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AgreementsConditionalChargesMPSSelect_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [ConditionalChargesMPSSelect$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sMenuItem = adoData( "MenuItem")  & ""

	'Flow

	Browser("index:=0").Page("index:=0").Sync

    Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Renewal Escalator").SiebApplet("Line Items")


	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
		End If

		If sLocateCol <> "" Then
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
			If  Cstr(res)<>"True" Then
				 iModuleFail = 1
				AddVerificationInfoToResult  "Error" ,sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

		adoData.MoveNext
	Loop

	SiebApplication("Siebel Call Center").SiebScreen("Agreements").SiebView("Renewal Escalator").SiebApplet("Line Items").SiebMenu("Menu").SiebSelectMenu sMenuItem


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : HandlePopup_fn
' 	Description :  This function is used to handle popups
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function HandlePopup(sPopUp)

		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

	'sOR
'	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Set vDictionary=CreateObject("Scripting.Dictionary") 
		vResult=CheckForPopup(vDictionary,sPopUp)

		If sPopUp <> False and not(instr(sPopUp,"OPTIONAL:")>0) Then
			expPopUp = True
			If vResult <> expPopUp Then
				AddVerificationInfoToResult "Fail", "Expected pop up - " & sPopUp & "did not occur"
				iModuleFail = 1
				ExitAction("FAIL")
			End If
		End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : HandlePopup_fn
' 	Description :  This function is used to handle popups
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function HandlePopup_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [PopUp$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

	'sOR
'	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	Set vDictionary=CreateObject("Scripting.Dictionary") 
		vResult=CheckForPopup(vDictionary,sPopUp)

		If sPopUp <> False and not(instr(sPopUp,"OPTIONAL:")>0) Then
			expPopUp = True
			If vResult <> expPopUp Then
				AddVerificationInfoToResult "Fail", "Expected pop up - " & sPopUp & "did not occur"
				iModuleFail = 1
				ExitAction("FAIL")
			End If
		End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : ServiceRequestLists_fn
' 	Description :  This function is used to click on new button in Service accounts applet on account s page.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ServiceRequestLists_fn()


	Dim adoData	  
	strSQL = "SELECT * FROM [ServiceRequestList$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

	sRequestSource = adoData( "RequestSource")  & ""
	sExpectedSource = adoData( "ExpectedSource")  & ""
	sNormalValidation = adoData( "NormalValidation")  & ""
	sClickSRResolved = adoData( "ClickSRResolved")  & ""
	sClickSubmitSR = adoData( "ClickSubmitSR")  & ""
	sType = adoData( "Type")  & ""
	sArea = adoData( "Area")  & ""
	sSubArea = adoData( "SubArea")  & ""
	sResolutionCode = adoData( "ResolutionCode")  & ""
	sEnter = adoData( "Enter")  & ""
	sGetStatus = adoData( "GetStatus")  & ""
	sEnterSRResoultionCode = adoData( "EnterSRResoultionCode")  & ""

	'Flow


	Browser("index:=0").Page("index:=0").Sync

	If sNormalValidation = "Yes" Then
	
		If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Request Source").Exist(10) Then
			SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Request Source").SiebSelect sRequestSource
		Else
			AddVerificationInfoToResult  "Error" , "Service Request List page is not displayed successfully."
			iModuleFail = 1	
		End If
	
		sSelectedSource = Trim(SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Request Source").GetRoProperty("activeitem"))
	
		If sSelectedSource <> sExpectedSource Then
			AddVerificationInfoToResult  "Info" , "User is able to edit Request source weblist"
		Else
			AddVerificationInfoToResult  "Error" , "User isnot  able to edit Request source weblist"
			iModuleFail = 1
		End If
	
	End If

	If sClickSRResolved = "Yes" AND sClickSubmitSR = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("SR Resolved").SiebClick sPopUp
		On Error Resume Next
		err.clear
	End If

	If sEnter = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Type").SiebSelect sType
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Area").SiebSelect sArea
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Sub area").SiebSelect sSubArea
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("SR Resolution Code").SiebSelect sResolutionCode
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebTextArea("Resolution Comment").SetText "abc"
	End If

	If sEnterSRResoultionCode = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("SR Resolution Code").SiebSelect sResolutionCode
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebTextArea("Resolution Comment").SetText "abc"
	End If

	If sClickSubmitSR = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("Submit SR").SiebClick False
	End If	


	If sClickSRResolved = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebButton("SR Resolved").SiebClick False
	End If

	If sGetStatus = "Yes" Then

		sStatus = Trim(SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("Service Request Activities").SiebApplet("Service Requests").SiebPicklist("Status").GetRoProperty("activeitem"))
		If sStatus = "Closed" Then
			AddVerificationInfoToResult  "Info" ,"SR status is closed"
		Else
			AddVerificationInfoToResult  "Error" ,"SR status is : " & sStatus
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
' 	Function Name : ServiceRequestAccountSelection_fn
' 	Description :  This function is used to select account number from service request screen
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function ServiceRequestAccountSelection_fn()


	Dim adoData	  
	strSQL = "SELECT * FROM [ServiceRequestAccountSelection$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sVal = adoData( "Value")  & ""
	sFind = adoData( "Find")  & ""
	sStartingWith = adoData( "StartingWith")  & ""
	
	sAccountNumber = DictionaryTest_G.Item("AccountNo")

	If sStartingWith = "" Then
		sStartingWith = sAccountNumber
	End If

	'Flow

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("classname:=SiebApplet","repositoryname:=SR " & sVal & " Pick Applet").Exist(10) Then
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("classname:=SiebApplet","repositoryname:=SR " & sVal & " Pick Applet").SiebPicklist("classname:=SiebPicklist","repositoryname:=PopupQueryCombobox").SiebSelect sFind
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("classname:=SiebApplet","repositoryname:=SR " & sVal & " Pick Applet").SiebText("classname:=SiebText","repositoryname:=PopupQuerySrchspec").SiebSetText sStartingWith
		SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("classname:=SiebApplet","repositoryname:=SR " & sVal & " Pick Applet").SiebButton("classname:=SiebButton","repositoryname:=PopupQueryExecute").SiebClick False
		If SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("classname:=SiebApplet","repositoryname:=SR " & sVal & " Pick Applet").SiebButton("classname:=SiebButton","repositoryname:=PickRecord").Exist(2) Then
			 SiebApplication("Siebel Call Center").SiebScreen("Service Requests").SiebView("My Service Requests").SiebApplet("classname:=SiebApplet","repositoryname:=SR " & sVal & " Pick Applet").SiebButton("classname:=SiebButton","repositoryname:=PickRecord").SiebClick False
		End If
	Else
		AddVerificationInfoToResult  "Error" , "Open pop up is not clicked successfuly for" & sVal & "  in Service request tab."
		iModuleFail = 1
	End If


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : CreateNewProfilesBillingProfile_fn
' 	Description :  This function is used to click on Profile tab and create new billing profile.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CreateNewProfilesBillingProfile_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [ProfilesBillingProfile$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	sClickCardDetails = adoData( "CardDetails")  & ""
	sNew = adoData( "New")  & ""
	sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Com Account Invoice Profile View (Invoice)","L3"
	

	If sNew = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebButton("New").SiebClick False
	End If

	If  sNew = "NewButton" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebButton("New").SiebClick False
	End If
	


	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebList("List").ActivateRow 0

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
				AddVerificationInfoToResult  "Error" ,sLocateCol & "-" & sLocateColValue & " not found in the list."
				Exit Function
			End If
		End If

		If sUIName <> "" Then		
			UpdateSiebList objApplet,sUIName,sValue
		End If

		If sLocateColValue = "Direct Debit" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebMenu("Menu").Select "ExecuteQuery"
'			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Billing profile").SiebMenu("Menu").Select "AboutRecord"
		End If
		adoData.MoveNext
	Loop

	If sNew = "RealTimeBalance" Then
		If SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Product / Services").SiebButton("Top-Up Request").Exist(5) Then
			AddVerificationInfoToResult  "Info" ,"Real Time Balance page is displayed directly after drill down on Name column in Profiles Billing Profile applet."
			CaptureSnapshot()
			Exit Function
		Else
			AddVerificationInfoToResult  "Error" ,"Real Time Balance page is not displayed directly after drill down on Name column in Profiles Billing Profile applet."
			iModuleFail = 1
			Exit Function
		End If
	End If

	If sClickCardDetails <> "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Credit/Debit Card Payment").SiebButton("Card Details").SiebClick sPopUp

	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : AddressesDeleteButtonCheck_fn
' 	Description :  This function is used to check whether Delete button is enabled or disabled at addresses tab
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AddressesDeleteButtonCheck_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [AddressesDeleteButton$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	sUser = adoData( "User")  & ""

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "Account Detail - Business Address View","L3"

	Select Case sUser

		Case "TEST_RETAIL_1"

			If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("Delete").Exist(5) Then
				AddVerificationInfoToResult  "Info" , "Delete button is disabled for User : " & sUser
			Else
				AddVerificationInfoToResult  "Error" , "Delete buuton is not enabled for User : " & sUser
				iModuleFail = 1
			End If

			CaptureSnapshot()

		Case "TEST_FRS"

			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").SiebButton("Delete").Exist(5) Then
				AddVerificationInfoToResult  "Info" , "Delete button is enabled for User : " & sUser
			Else
				AddVerificationInfoToResult  "Error" , "Delete button is disabled for User : " & sUser
				iModuleFail = 1
			End If

			CaptureSnapshot()

		Case "FIXED_LINE_1"

			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Address Book").SiebApplet("Account Addresses").SiebButton("Delete").Exist(5) Then
				AddVerificationInfoToResult  "Info" , "Delete button is enabled for User : " & sUser
			Else
				AddVerificationInfoToResult  "Error" , "Delete button is disabled for User : " & sUser
				iModuleFail = 1
			End If

			CaptureSnapshot()

	End Select
	

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function






'#################################################################################################
' 	Function Name : MonetaryBalanceCheck_fn
' 	Description : This function is used to click on Real Time Balance
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function MonetaryBalanceCheck_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [MonetaryBalance$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	sSelectItem = adoData( "SelectItem")  & ""
	sPart = adoData( "Part")  & ""
    sCheckProductName = adoData( "CheckProductName")  & ""
	sClickButton = adoData( "ClickButton")  & ""

	If DictionaryTest_G.Exists(sSelectItem) Then
			sSelectItem = Replace(sSelectItem,sSelectItem,DictionaryTest_G(sSelectItem))
	
	elseif DictionaryTempTest_G.Exists(sSelectItem) Then
			sSelectItem = Replace(sSelectItem,sSelectItem,DictionaryTempTest_G(sSelectItem))
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

	If  sClickButton= "FullBalance" Then '' Click on Full Balance button
		SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebButton("Full Balance").SiebClick False
	End If
	
	If   sClickButton= "MonetaryBalance" Then ''MonetaryBalance Click
		SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebButton("Monetary Balance").SiebClick False
	End If



	If  Ucase(sCheckProductName) = "YES" Then


			Do while Not adoData.Eof
							sLocateCol = adoData( "LocateCol")  & ""
							sLocateColValue = adoData( "LocateColValue")  & ""
							sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
							sExist = adoData( "Exist")  & ""
						If sLocateCol <> "" Then
							res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
							
							If sExist = "" Then
								If  res=True Then
									AddVerificationInfoToResult  "Pass" , sLocateCol & "-" & sLocateColValue & " found in the list as expected"				 
					'				Exit Function
								else
									iModuleFail = 1
									AddVerificationInfoToResult  "Fail" , sLocateCol & "-" & sLocateColValue & " not found in the list."
									
								End If 
							else
								If sExist = "False" Then
									sExist = "False-Row Not Exist"
								End If
								If Ucase(cstr(res)) = Ucase (sExist) Then
									AddVerificationInfoToResult  "Pass" , sLocateCol & "-" &  sLocateColValue & " existence is " & sExist & " as expected"
								else
									 iModuleFail = 1
									AddVerificationInfoToResult  "Fail" , sLocateCol & "-" & sLocateColValue & " existence is " & cstr(res) & " but expected is " & sExist
								End If
							End If
						End If
						CaptureSnapshot()
						adoData.MoveNext
			Loop
							
	
		Else
					If sPart = "First" Then	
						sRemainingQuantity = Trim(SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebList("List").SiebGetCellText("Remaining qty",0))
				
						If DictionaryTest_G.Exists("MonetaryBalance") Then
							DictionaryTest_G.Item("MonetaryBalance")=sRemainingQuantity
						else
							DictionaryTest_G.add "MonetaryBalance",sRemainingQuantity
						End If
					
						CaptureSnapshot()
					
						AddVerificationInfoToResult  "Info" , "Initial Balance before adding bundle  is " & sRemainingQuantity
				
				'        SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills_2").SiebApplet("Billing profile").SiebButton("Customer Account").SiebClick False
				
						  SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Billing profile").SiebButton("Customer Account").SiebClick False
				
				
					ElseIf sPart = "Second" Then
						sBundleAddition = SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Balance Details").SiebList("List").SiebGetCellText("Remaining qty",0)
				
						If Trim(sBundleAddition) <> DictionaryTest_G.Item("MonetaryBalance") Then
							AddVerificationInfoToResult  "Info" , "Balance is changed to " & sBundleAddition & " after adding bundle"
						Else
							AddVerificationInfoToResult  "Info" , "Balance is not changed after adding bundle"
							iModuleFail = 1
						End If
						CaptureSnapshot()
					End If
						
	End If
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : CustomerCommsButtonsValidation_fn
' 	Description : This function is used to click on customer comms page and validate Set/Reset Password, Pin and Word-Hint buttons
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CustomerCommsButtonsValidation_fn()
	
	'sOR
	call SetObjRepository ("Catalog",sProductDir & "Resources\")

	'Flow

	 SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Activities Screen"  ' Click on Customer Comms tab
     'SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Customer Comms").SiebButton("Set/Reset Password").Click
     'SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Customer Comms").SiebButton("Set/Reset PIN").Click 
     'SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Customer Comms").SiebButton("Set/Reset Word and Hint").Click


	If SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Customer Comms").Exist(5) Then
		If ((Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("innertext:=Set/Reset Password","index:=0","class:=miniBtnUICOff").Exist(3)) OR (SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Customer Comms").SiebButton("Set/Reset Password").exist(2))) Then
			AddVerificationInfoToResult  "Info" , "Set/Reset Password button is present on page."
		Else
			AddVerificationInfoToResult  "Error" , "Set/Reset Password button is not present on page."
			iModuleFail = 1
		End If

		If ((Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("innertext:=Set/Reset PIN","index:=0","class:=miniBtnUICOff").Exist(3)) OR (SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Customer Comms").SiebButton("Set/Reset PIN").Exist(2)))Then
			AddVerificationInfoToResult  "Info" , "Set/Reset PIN button is present on page."
		Else
			AddVerificationInfoToResult  "Error" , "Set/Reset PIN button is not present on page."
			iModuleFail = 1
		End If

		If ((Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("innertext:=Set/Reset Word and Hint","index:=0","class:=miniBtnUICOff").Exist(3)) OR(SiebApplication("Siebel Call Center").SiebScreen("Customer Comms").SiebView("My Activities").SiebApplet("Customer Comms").SiebButton("Set/Reset Word and Hint").Exist(2)))Then
			AddVerificationInfoToResult  "Info" , "Set/Reset Word and Hint button is present on page."
		Else
			AddVerificationInfoToResult  "Error" , "Set/Reset Word and Hint button is not present on page."
			iModuleFail = 1
		End If
	Else
		AddVerificationInfoToResult  "Error" , "Customer Comms Tab is not clicked successflly."
		iModuleFail = 1
		Exit Function
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : Profiles_DirectDebit_fn
' 	Description : This function is used to click on customer comms page and validate Set/Reset Password, Pin and Word-Hint buttons
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Profiles_DirectDebit_fn()
	
	'sOR
	call SetObjRepository ("Catalog",sProductDir & "Resources\")

	'Flow
	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebPicklist("DD Transaction Type").Exist(5) Then
		sDDTransactionType = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Direct Debit").SiebPicklist("DD Transaction Type").GetRoProperty("activeitem")
		If Instr(sDDTransactionType,"0C") > 0 Then
			AddVerificationInfoToResult  "Info" ,"DD Transaction Type value is changed to " & sDDTransactionType & " after changing payment type fron DD to Bill Me."
		Else
			AddVerificationInfoToResult  "Error" ,"DD Transaction Type value is not changed to 0C after changing payment type fron DD to Bill Me."
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
' 	Function Name : OrderBillingServiceAccountVerify_fn
' 	Description :  This function is used to Verify  billing account and service account from orders page.
'   Created By :  Pushkar
'	Creation Date :        
'##################################################################################################
Public Function OrderBillingServiceAccountVerify_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [BillingServiceAccount$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Catalog.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sAccountSelection = adoData( "AccountSelection")  & ""
	sAccountSet = adoData( "AccountSet")  & ""

	Browser("index:=0").Page("index:=0").Sync

	If sAccountSelection = "Billing Account" Then
		   If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").Exist(2) Then
				sBillAccountName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").GetROProperty("text")
			Else
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").click
				sBillAccountName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing account").GetROProperty("text")
			End If
				If Ucase(sAccountSet) = Ucase(sBillAccountName) Then
					AddVerificationInfoToResult  "Pass" , "Billing Account is reflected successfully."
				Else
						iModuleFail = 1
						AddVerificationInfoToResult  "Fail" ,"Billing Account is not reflected successfully."
				End If

	End If
					
			CaptureSnapshot()
			
		If sAccountSelection = "Service Account" Then
			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").Exist(2) Then
				sServiceAccountName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").GetROProperty("text")
			Else
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").click
				sServiceAccountName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").GetROProperty("text")
			End If
					If Ucase(sAccountSet) = Ucase(sServiceAccountName) Then
					AddVerificationInfoToResult  "Pass" , "Service  Account is reflected successfully."
					Else
					iModuleFail = 1
					AddVerificationInfoToResult  "Fail" ,"Service Account is not reflected successfully."
				End If
	End If
		
		CaptureSnapshot()

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : Bills_EmailCopy_fn
' 	Description : This function is used to click on customer comms page and validate Set/Reset Password, Pin and Word-Hint buttons
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Bills_EmailCopy_fn()
	
	'sOR
	call SetObjRepository ("Catalog",sProductDir & "Resources\")

	'Flow
	If SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Bills").Exist(5) Then
		AddVerificationInfoToResult  "Info" , "Name link is drilled down successfully."
		SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Bills").SiebButton("Email Copy Bill").Click
		If SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Select Email Address").Exist(5) Then
			AddVerificationInfoToResult  "Info" , "Email Copy Bill buttion is clicked successfully."
			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Select Email Address").SiebPicklist("Select Email address").SiebSelect "Account holder Email"
			CaptureSnapshot()
			SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Select Email Address").SiebButton("Send Email").Click
			If Browser("Siebel Call Center").Dialog("Siebel").Static("Your request has been").Exist(5) Then
				AddVerificationInfoToResult  "Info" , "Pop up message is displayed successfully."
				CaptureSnapshot()
				Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
			Else
				AddVerificationInfoToResult  "Error" , "Pop up message is not displayed successfully."
				iModuleFail = 1
			End If
		Else
			AddVerificationInfoToResult  "Error" , "Email Copy Bill is not clicked successfully."
			iMModuleFail = 1
			Exit Function
		End If
	Else
		AddVerificationInfoToResult  "Error" , "Account Name is not drilled down in Profile tab."
		Exit Function
		iModuleFail = 1
	End If

	SiebApplication("Siebel Call Center").SiebScreen("CMU Billing Profile Portal").SiebView("Bills").SiebApplet("Billing profile").SiebButton("Customer Account").SiebClick False

	Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function
