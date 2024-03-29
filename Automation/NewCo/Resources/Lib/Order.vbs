'#################################################################################################
' 	Function Name : OrdersPayment_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrdersPayment_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [Order_Payment$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow
	sMSISDN = DictionaryTest_G("ACCNTMSISDN")
		Browser("index:=0").Page("index:=0").Sync


'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCheckbox("Terms and conditions").SiebSetCheckbox "ON"

'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Payment View (Sales)","L3"

		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview") <> "Order Entry - Payment View (Sales)" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Payment View (Sales)","L3"
		End If

		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview") <> "Order Entry - Payment View (Sales)" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Payment View (Sales)","L3"
		End If



		Do while Not adoData.Eof

		sPaymentMethod = adoData( "PaymentMethod")  & ""
		sTransactionType = adoData( "TransactionType")  & ""
		sTransactionAmount = adoData("TransactionAmount") & ""
		sJourneyType = adoData("JourneyType") & ""
		sAuthorize = adoData("Authorize") & ""
		sCVV = adoData("CVV") & ""
		sDeletePayment = adoData( "DeletePayment")  & ""
		sClickCardDetails = adoData( "ClickCardDetails")  & ""
		sClickCardDetails1 = adoData( "ClickCardDetails1")  & ""
		sClickNew = adoData( "ClickNew")  & ""

		sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

		sPopUp1 = adoData( "Popup1")  & ""	
		If Ucase(sPopUp1)="NO" Then
			sPopUp1="FALSE"
			sPopUp1=Cbool(sPopUp1)
		End If

		sPopUp2 = adoData( "PopUp2")  & ""	
		If Ucase(sPopUp2)="NO" Then
			sPopUp2="FALSE"
			sPopUp2=Cbool(sPopUp2)
		End If
	
			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").Exist(3) Then
				If sClickNew <> "No" Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebButton("New").SiebClick sPopUp ' Clicking on New Button
				End If
			Else
				AddVerificationInfoToResult "Fail","Payments tab is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If
	
			
			rowCount = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").RowsCount 'Taking Row count after clicking on New Button
	
			If  rowCount <> 0 Then
				If sPaymentMethod<>"" Then			
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").SiebPicklist("Payment method").SiebSelectPopup sPaymentMethod,sPopUp2 ' Setting Payment Method
				End If

				If sPaymentMethod="Credit/Debit Card"  AND sAuthorize = "Yes" Then
                	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").SiebText("Billing profile").OpenPopup
					If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").Exist(15) Then
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebPicklist("SiebPicklist").Select "Payment method"
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebText("SiebText").SiebSetText sPaymentMethod
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebButton("Go").SiebClick False
'						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebButton("OK").SiebClick False
					Else
						AddVerificationInfoToResult "Fail","Billing Profile window is not displayed after clicking onPop up button.."
						iModuleFail = 1
					End If
				End If

					If sPaymentMethod="On Account"  AND sJourneyType = "Postpaid" Then
                	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").SiebText("Billing profile").OpenPopup
					If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").Exist(15) Then
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebPicklist("SiebPicklist").Select "Postpay/ Prepay"
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebText("SiebText").SiebSetText sJourneyType
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebButton("Go").SiebClick False
'						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebButton("OK").SiebClick False
					Else
						AddVerificationInfoToResult "Fail","Billing Profile window is not displayed after clicking onPop up button.."
						iModuleFail = 1
					End If
				End If

				If sPaymentMethod="On Account"  AND sJourneyType = "Prepaid" Then
                	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").SiebText("Billing profile").OpenPopup
					If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").Exist(15) Then
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebPicklist("SiebPicklist").Select "Postpay/ Prepay"
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebText("SiebText").SiebSetText sJourneyType
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebButton("Go").SiebClick False
'						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Billing profile").SiebButton("OK").SiebClick False
					Else
						AddVerificationInfoToResult "Fail","Billing Profile window is not displayed after clicking onPop up button.."
						iModuleFail = 1
					End If
				End If

				If sPaymentMethod = "Balance" Then
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Detail - Balance").SiebText("Mobile no.").OpenPopup
						Browser("index:=0").Page("index:=0").Sync

						If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Pick Asset").Exist(5) Then
							CaptureSnapshot()
							SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Pick Asset").SiebText("SiebText").SiebSetText sMSISDN
							SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Pick Asset").SiebButton("Go").Click
							AddVerificationInfoToResult "Info","MSIDNchosen : "&sMSISDN
'							SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Pick Asset").SiebButton("OK").Click
						End If
						Browser("index:=0").Page("index:=0").Sync
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Detail - Balance").SiebButton("Reserve Funds").SiebClick False
						CaptureSnapshot()
				End If

				If sClickCardDetails = "Yes" Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Details - Credit").SiebButton("Card Details").SiebClick sPopUp1
				End If


				If sTransactionAmount<>"" Then			
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").SiebCalculator("classname:=SiebCalculator","repositoryname:=Transaction Amount").setText sTransactionAmount' Setting Payment Amount
				End If
	
				If sTransactionType<>"" Then			
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").SiebPicklist("Transaction type").SiebSelect sTransactionType ' Setting Transaction Type
				End If

				If sClickCardDetails1 = "Yes" Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Details - Credit").SiebButton("Card Details").SiebClick sPopUp
				End If

				If sAuthorize = "Yes" Then
					If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Details - Credit").SiebButton("Authorise").getRoProperty("IsEnabled") Then
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Details - Credit").SiebButton("Authorise").SiebClick False
						If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("CV2").Exist(5) Then
							SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("CV2").SiebText("CV2").SiebSetText sCVV
							SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("CV2").SiebButton("Go").SiebClick False
						Else
							AddVerificationInfoToResult "Fail","CVV window is not displayed after clicking on Authorise button."
							iModuleFail = 1
							Exit Function
						End If
						
						If sDeletePayment = "Y" Then
							SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebButton("Delete").SiebClick False

							wait 2

							rwCnt = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebList("List").SiebListRowCount

							If rwCnt = 0 Then
								AddVerificationInfoToResult "Pass","Agent User is able to delete an Authorized credit card payment line."
							Else
								AddVerificationInfoToResult "Fail","Agent User is not able to delete an Authorized credit card payment line."
								iModuleFail = 1
							End If
						End If
					Else
						AddVerificationInfoToResult "Fail","Authorize button is not enabled."
						iModuleFail = 1
						Exit Function
					End If
				End If
'				AddVerificationInfoToResult "Order Payment - Pass","New Button is clicked and Payment method is selected successfully."
			Else
				AddVerificationInfoToResult "Order Payment - Fail","New Button is not clicked successfully."
				iModuleFail = 1
				Exit Function
			End If
			adoData.movenext
		Loop

		If sJourneyType = "INC" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebMenu("Menu").Select "ExecuteQuery"

				If Browser("Siebel Call Center").Dialog("Siebel").Static("As you have selected On").Exist(5) Then
					CaptureSnapshot()
					Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
					AddVerificationInfoToResult  "Info" ,"Pop up message is displayed successfully when you have selected OnAccount payment method"
				Else
					AddVerificationInfoToResult  "Error" ,"Pop up message is not displayed successfully when you have selected OnAccount payment method."
					iModuleFail = 1
				End If
		End If

		If sJourneyType = "MaxAmount" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Lines").SiebMenu("Menu").Select "ExecuteQuery"

				If Browser("Siebel Call Center").Dialog("Siebel").Static("The maximum amount for").Exist(5) Then
					wait 2
					CaptureSnapshot()
'					Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
					AddVerificationInfoToResult  "Info" ,"Pop up message is displayed successfully when transaction amount exceeds the maximum limit."
				Else
					AddVerificationInfoToResult  "Error" ,"Pop up message is not displayed successfully when transaction amount exceeds the maximum limit."
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
' 	Function Name : OrdersLineItemsVerify_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################

Public Function OrdersLineItemsVerify_fn()

	Dim rwNum
	Dim loopVar
	Dim strActionReposName
	Dim strProductReposName
	Dim strInstalledIdReposName
	Dim strProduct
	Dim strAction
	Dim strActionDataSheet
	Dim sProduct
	Dim adoData
	Dim strSQL

		  
    strSQL = "SELECT * FROM [LineItems$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")


	sExpand = adoData( "Expand")  & ""


	sVerifyOutcome= adoData( "VerifyOutcome")  & ""


	Browser("index:=0").Page("index:=0").Sync		
	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If
''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************

	CaptureSnapshot()
	Browser("index:=0").Page("index:=0").Sync	

	If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").Exist Then
		AddVerificationInfoToResult "Info","Items are displayed successfully."
	Else
		iModuleFail = 1
		AddVerificationInfoToResult "Fail","Items are not displayed successfully."
	End If
	
'	If sExpand<>"" Then
'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectExpandRow sExpand,0
'	End If


	Call ShowMoreButton_fn()
	Browser("index:=0").Page("index:=0").Sync
	call SetObjRepository ("Orders",sProductDir & "Resources\")
	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items")

	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
		'sLocateColValue = Replace(sLocateColValue,"RootProduct0",DictionaryTest_G.Item("ROOT_PRODUCT"))
		sExist = adoData( "Exist")  & ""
		sExpand = adoData( "Expand")  & ""
		'sExpand = Replace(sExpand,sExpand,DictionaryTest_G.Item("ROOT_PRODUCT"))
		If DictionaryTest_G.Exists(sExpand) Then
'			sExpand = Replace(sExpand,sExpand,DictionaryTest_G.Item("ROOT_PRODUCT"))
			sExpand = Replace(sExpand,sExpand,DictionaryTest_G(sExpand))
		elseif DictionaryTempTest_G.Exists(sExpand) Then
			sExpand = Replace(sExpand,sExpand,DictionaryTempTest_G(sExpand))
		End If 
		sCollapse = adoData( "Collapse")  & ""

		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
'		else
'			iIndex=Cint(iIndex)
		End If

		If sExpand<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectExpandRow sExpand,iIndex
		End If

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
		If sCollapse<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebMenu("Menu").SiebSelectMenu "Run Query"
'			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectCollapseRow sCollapse,iIndex
		End If
		adoData.MoveNext
	Loop
'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectCollapseRow sExpand,0
	If (sVerifyOutcome <> "ServiceAccountLastName" AND sVerifyOutcome <> "")Then
		sOutcomeActual = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Credit Vetting Results").SiebText("Outcome").GetROProperty("text")
			If Trim(sOutcomeActual) = Trim(sVerifyOutcome) Then
				AddVerificationInfoToResult "Info" , "Outcome is " &sOutcomeActual
			Else
				AddVerificationInfoToResult "Fail" , "Outcome is" &sOutcomeActual
				iModuleFail = 1
			End If
	End If


		If  sVerifyOutcome = "ServiceAccountLastName" Then
					strEnable = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebText("User account").getROProperty("isenabled")
					CaptureSnapshot
					If  strEnable = True Then
								AddVerificationInfoToResult "Pass","Service Account is editable.."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Service Account is not editable."
						Exit Function
						End If
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebText("User account").OpenPopup
						If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").exist(2)Then
							rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebList("List").SiebListRowCount
							CaptureSnapshot
									If rowCnt <> 0 Then
										AddVerificationInfoToResult "Pass","Rows are displayed. Service Account is editable.."
									Else
										iModuleFail = 1
										AddVerificationInfoToResult "Fail","No rows displayed. Service Account is not editable."
										Exit Function
									End If
								SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Click				
						End If
							

						strEnable = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebText("Last name").getROProperty("isenabled")
						CaptureSnapshot
						If  strEnable = True Then
								AddVerificationInfoToResult "Pass","Last Name is editable.."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Last Name is not editable."
						Exit Function
						End If
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebText("Last name").OpenPopup
						If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Contact").exist(2)Then
							rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Contact").SiebList("List").SiebListRowCount
							CaptureSnapshot

									If rowCnt <> 0 Then
										AddVerificationInfoToResult "Pass","Rows are displayed. Last Name is editable.."
									Else
										iModuleFail = 1
										AddVerificationInfoToResult "Fail","No rows displayed. Last Name is not editable."
										Exit Function
									End If
								SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Contact").SiebButton("OK").Click				
						
						End If
			End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function






'#################################################################################################
' 	Function Name : OrdersLineItemsEntry_fn()
' 	Description : This function is used to search the Account number
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################

Public Function OrdersLineItemsEntry_fn()

	Dim rwNum
	Dim loopVar
	Dim strActionReposName
	Dim strProductReposName
	Dim strInstalledIdReposName
	Dim strProduct
	Dim strAction
	Dim strActionDataSheet
	Dim sProduct
	Dim adoData
	Dim strSQL

		  
    strSQL = "SELECT * FROM [OrdersLineItemsEntry$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")



	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If	
''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************
	sChangeMSISDN = adoData( "ChangeMSISDN")  & ""
	sTVDiscount = adoData( "TVDiscount")  & ""
	

'	If sChangeMSISDN<>"" Then
'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebText("Installed ID").OpenPopup
'        CaptureSnapshot()
'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Resource").SiebButton("OK").SiebClick False
'		AddVerificationInfoToResult  "Info" , "MSISDN is selected from picklist."
'		Exit Function
'	End If


	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items")
'	Call ShowMoreButton_fn()
	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
'		sLocateColValue = Replace(sLocateColValue,"RootProduct0",DictionaryTest_G.Item("ROOT_PRODUCT"))
        sLocateColValue = Replace(sLocateColValue,"MSISDN",DictionaryTest_G.Item("ACCNTMSISDN"))
		sUIName = adoData( "UIName")  & ""
		sValue = adoData( "Value")  & ""
		sExpand = adoData( "Expand")  & ""
		'sExpand = Replace(sExpands,sExpand,DictionaryTest_G.Item("ROOT_PRODUCT"))
		If DictionaryTest_G.Exists(sExpand) Then
			'sExpand = Replace(sExpands,sExpand,DictionaryTest_G.Item("ROOT_PRODUCT"))
			sExpand = Replace(sExpand,sExpand,DictionaryTest_G(sExpand))
		elseif DictionaryTempTest_G.Exists(sExpand) Then
'			sExpand = Replace(sExpand,sExpand,DictionaryTempTest_G.Item("ROOT_PRODUCT"))
		sExpand = Replace(sExpand,sExpand,DictionaryTempTest_G(sExpand))
		End If 
		sCollapse = adoData( "Collapse")  & ""
		sAction = adoData( "Action")  & ""
		sKey = adoData( "Key")  & ""
		Dim key1
		If instr(sUIName,"CaptureTextValue") > 0 and sValue <> "" Then
			key1 = sValue
		End If

		If sValue="ICCID" Then ' ICC Id
			sValue=DictionaryTest_G.Item("ICCID")
		End If

		If sValue="RECONNECTIONICCID" Then ' ICC Id
			sValue=DictionaryTest_G.Item("RECONNECTIONICCID")
		End If
	
		If sValue="ROUTER NUMBER" Then ' Router Number	
			sValue=DictionaryTest_G.Item("ROUTER")
		End If

		If sValue="MSISDN" Then '  Account MSISDN for Reconnection or Manual
			sValue=DictionaryTest_G.Item("ACCNTMSISDN")
		End If

		If sValue="NEWACCNTMSISDN" Then '  New Account MSISDN for Reconnection or Manual
			sValue=DictionaryTest_G.Item("NEWACCNTMSISDN")
		End If

		If sValue= "BBIPNUM" Then '  Account MSISDN
			sValue = RetrieveBBIP
		End If

		If sValue="STB NUMBER" Then '  STB number for TV
			If DictionaryTest_G.Exists(sKey) Then
				sValue=DictionaryTest_G.Item(sKey)
			elseif DictionaryTempTest_G.Exists(sKey) Then
				sValue=DictionaryTempTest_G.Item(sKey)
			End If
		End If
'
'		If sValue="STB NUMBER" Then '  STB number for TV
'			sValue=DictionaryTest_G.Item(sKey)
'		End If 

		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
		End If

		If sExpand<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectExpandRow sExpand,iIndex
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

		If instr(sUIName,"CaptureTextValue") > 0 Then
			DictionaryTest_G (key1) = sValue
		End If

			If instr(sUIName,"CaptureTextValue") > 0 And instr(sValue,"RECONNECTIONICCID") > 0  Then
			DictionaryTest_G.Item("RECONNECTIONICCID") = sValue
		End If

	

		CaptureSnapshot()
		adoData.MoveNext

        If sChangeMSISDN = "Y" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Resource").Exist(5) Then
			CaptureSnapshot()
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Resource").SiebButton("OK").SiebClick False
			AddVerificationInfoToResult  "Info" , "MSISDN is selected from picklist."
		Else
			AddVerificationInfoToResult  "Error" , "MSISDN is not selected from picklist."
			iModuleFail = 1
		End If
	End If

	If (sChangeMSISDN="N" )Then		
				If  Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebElement("OK").Exist(5) Then
						 AddVerificationInfoToResult  "PASS" , "OK Button is Disabled as Expected in MSISDN Popup"
						Browser("SiebWebPopupWindow").close
						Exit Function
				Else
						 iModuleFail = 1
						AddVerificationInfoToResult  "FAIL" , "OK Button is Enabled."
						Exit Function
				End If
	End If

		If sAction = "Reconnection" Then
			If  Browser("Siebel Call Center").Dialog("Siebel").Exist(2) Then
				Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
		
			End If
		End If


		If sCollapse<>"" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebMenu("Menu").SiebSelectMenu "Run Query"
		End If
	Loop

		If sTVDiscount <> "" Then		
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebPicklist("% Discount").SiebSelect "100%"
		End If

    	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function


'#################################################################################################
' 	Function Name : MenuLineItems_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function MenuLineItems_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [MenuLineItems$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sMenuItem = adoData( "MenuItem")  & ""

	Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"

		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview") <> "Order Entry - Line Items Detail View (Sales)" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
		End If

	Browser("index:=0").Page("index:=0").Sync
'''*****************Web Element********************
''Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
''Browser("index:=0").Page("index:=0").Sync
'''*****************Web Element********************
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCheckbox("Terms and conditions").SiebSetCheckbox "ON"
'	    Reporter.Filter = 0

'	SiebApplication("Siebel Call Center").SetTimeOut(300)

		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items")

	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
		sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
		'sLocateColValue = Replace(sLocateColValue,"RootProduct0",DictionaryTest_G.Item("RootProduct0"))

		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
'		else
'			iIndex=Cint(iIndex)
		End If

		If sLocateCol <> "" Then
			res=LocateColumns2 (objApplet,sLocateCol,sLocateColValue,iIndex)		
				If  res=True Then
					AddVerificationInfoToResult  "Pass" , sLocateCol & "-" & sLocateColValue & " found in the list as expected"	
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebMenu("Menu").SiebSelectMenu sMenuItem	 
	'				Exit Function
				else
					iModuleFail = 1
					AddVerificationInfoToResult  "Fail" , sLocateCol & "-" & sLocateColValue & " not found in the list."
					
				End If 		
		End If
		
		adoData.MoveNext
	Loop
'	
CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'
''#################################################################################################
'' 	Function Name : Econfig_fnOLD
'' 	Description : This function log into Siebel Application
''   Created By :  Ankit
''	Creation Date :        
''##################################################################################################
'Public Function EconfigOLD_fn()
'
'Dim adoData	  
'    strSQL = "SELECT * FROM [Econfig$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Econfig.xls",strSQL)
'
'	'sOR
'	call SetObjRepository ("Orders",sProductDir & "Resources\")
'
'	'Flow
'		Do while Not adoData.Eof
'			sLink = adoData( "Link")  & ""
'			sItem = adoData( "Item")  & ""
'			sCheck = adoData( "Check")  & ""
'			sList = adoData( "List")  & ""
'			sCustomise= adoData( "Customise")  & ""
'			'sCustomiseItem= adoData( "CustomiseItem")  & ""
'			sTextValue= adoData( "TextValue")  & ""
'			sAddItem= adoData( "AddItem")  & ""
'			sAction = adoData( "Action")  & ""
'			sCatalog = adoData( "Catalog")  & ""
'
'			Browser("index:=0").Page("index:=0").Sync
'
'		    If  Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Exist(3) Then
'				 Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Click
'			End If
'
'			
'			If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Exist(5) Then
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Click
'			End If
'
'			Browser("index:=0").Page("index:=0").Sync
'
'			If sLink<>"" Then
'				If Instr(sLink,"|") > 0 Then
'					sLinkSplit = Split(sLink,"|")
'				End If
'				If Instr(sLink,"|") > 0 Then
'					For i = 0 to Ubound(sLinkSplit)
'
'						If sLinkSplit(i) = "RootProduct0" Then
'							If DictionaryTest_G.Exists(sLinkSplit(i)) Then
'									sLinkSplit(i) = Replace(sLinkSplit(i),sLinkSplit(i),DictionaryTest_G(sLinkSplit(i)))	
'								elseif DictionaryTempTest_G.Exists(sLinkSplit(i)) Then
'									sLinkSplit(i) = Replace(sLinkSplit(i),sLinkSplit(i),DictionaryTempTest_G(sLinkSplit(i)))
'								End If
'						End If
'	
'								If  sLinkSplit(i) = "Mobile phone service" Then
'										sLinkSplit(i) = Replace("Mobile phone service","Mobile phone service","Mobile Phone Services")	
'								End If
'		
'								If  sLinkSplit(i) = "Mobile broadband service" Then
'										sLinkSplit(i) = Replace("Mobile broadband service","Mobile broadband service","Mobile broadband Services")	
'										
'								End If
'	
'						Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("text:="&sLinkSplit(i),"html tag:=A").WebClick
'				
'							If sCheck<>"" Then
'								If sItem="A bar" or sItem="O bar" or sItem = "Credit Alert A Bar" or sItem="Fraud" or sItem="Stolen Bar" Then
'									sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").Exist(2)
'									If sExist = "True" Then
'										Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'									Else
'										iModuleFail = 1
'										Exit Function
'									End If
'								Else
'									sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("class:=eCfgSpan.*","innertext:="&sItem).Exist(2)
'									If sExist = "True" Then
'										AddVerificationInfoToResult "Info" , "Product : " & sItem & " is present in " & sLinkSplit(i) & " tab."
'										Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'										Exit For
'									End If
'								End If
'							End If
'					Next
'				Else
'						If sLink = "RootProduct0" Then
'							If DictionaryTest_G.Exists(sLink) Then
'									sLink = Replace(sLink,sLink,DictionaryTest_G(sLink))	
'								elseif DictionaryTempTest_G.Exists(sLink) Then
'									sLink = Replace(sLink,sLink,DictionaryTempTest_G(sLink))
'								End If
'						End If
'
'								
'		
'								If  sLink = "Mobile phone service" Then
'										sLink = Replace("Mobile phone service","Mobile phone service","Mobile Phone Services")
'								End If
'		
'								If  sLink = "Mobile broadband service" Then
'										sLink = Replace("Mobile broadband service","Mobile broadband service","Mobile broadband Services")
'								End If
'
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("text:="&sLink,"html tag:=A").WebClick
'
'						If sCheck<>"" Then
'							If sItem="A bar" or sItem="O bar" or sItem = "Credit Alert A Bar" or sItem="Fraud" or sItem="Stolen Bar" Then
'								sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").Exist(2)
'								If sExist = "True" Then
'									Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'								Else
'									AddVerificationInfoToResult "Error" , "Product : " & sItem & " is not present in " & sLink & " tab."
'									iModuleFail = 1
'									Exit Function
'								End If
'							Else
'								sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("class:=eCfgSpan.*","innertext:="&sItem).Exist(2)
'								If sExist = "True" Then
'									AddVerificationInfoToResult "Info" , "Product : " & sItem & " is present in " & sLink & " tab."
'									Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'								Else
'									AddVerificationInfoToResult "Error" , "Product : " & sItem & " is not present in " & sLink & " tab."
'									iModuleFail = 1
'									Exit Function
'								End If
'							End If	
'						End If
'				End If
'			End If
'
'
'			If sCatalog <> "" Then
'				If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("There is a conflict with").Exist(3) Then
'					AddVerificationInfoToResult "Info" , "Pop up message is displayed successfully after selecting " & sItem & " for Prepaid Account."
'					CaptureSnapshot()
'					Exit Function
'				Else
'					AddVerificationInfoToResult "Error" , "Pop up message is not displayed successfully."
'					iModuleFail = 1
'					Exit Function
'				End If
'			End If
'
'			If sCustomise="Y" Then
'				If sLink="Bars" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebTable("column names:=.*Customise").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
'				else
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
'				End If
'				
'				'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sCustomiseItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sCustomiseValue
'			End If			
'
'			If sList<>"" Then
'				If  sList="Any"Then
'					sList="#1"
'				End If
'				If sItem="Item" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sList&"']/..").WebSelect sList
'				ElseIf sItem ="AreaCode" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebList("all items:=.*"&sList&".*").WebSelect sList
'				Else
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//*[text()='"&sItem&"']/../..//select/option[text()='"&sList&"']/..").WebSelect sList			
'				End If
'			End If
'
'			If sAddItem<>"" Then
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").WebEdit("html tag:=INPUT").Set sAddItem
'
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").Link("text:=Add Item").WebClick
''				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT","index:=0").Set sAddItem
''				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Link("text:=Add Item").WebClick
'			End If	
'
'
'
''			If sTextValue<>"" Then
''				sTextValue = DictionaryTest_G.Item("ACCNTMSISDN")
'''                sTextValue = "447387945790"
''				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
''			End If
'
''*******Don not delete this code
'			If  sTextValue<>""Then
'				If sItem ="MSISDN" Then
'							sTextValue = DictionaryTest_G.Item("ACCNTMSISDN")
'			'                sTextValue = "447387945790"
'							Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
'				Else
'							Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
'			
'						End If
'
'			End If
''*******Don not delete this code
'
'
'			adoData.movenext
'		
'	Loop
'
'	CaptureSnapshot()
'
'	If Err.Number <> 0 Then
'			iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If
'
'
'End Function
'
'
''#################################################################################################
'' 	Function Name : Econfig_fn1
'' 	Description : This function log into Siebel Application
''   Created By :  Ankit
''	Creation Date :        
''##################################################################################################
'Public Function Econfig_fn1()
'
'Wait 5
'Dim adoData	 
'
'    strSQL = "SELECT * FROM [Econfig$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Econfig.xls",strSQL)
'
'	'sOR
'	call SetObjRepository ("Orders",sProductDir & "Resources\")
'
'	'Flow
'		Do while Not adoData.Eof
'			sLink = adoData( "Link")  & ""
'			sItem = adoData( "Item")  & ""
'
'			sCheck = adoData( "Check")  & ""
'
'			sList = adoData( "List")  & ""
'			sCustomise= adoData( "Customise")  & ""
'			'sCustomiseItem= adoData( "CustomiseItem")  & ""
'			sTextValue= adoData( "TextValue")  & ""
'			sAddItem= adoData( "AddItem")  & ""
'			sAction = adoData( "Action")  & ""
'			sCatalog = adoData( "Catalog")  & ""
'
'			Browser("index:=0").Page("index:=0").Sync
'
'		    If  Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Exist(3) Then
'				 Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Click
'			End If
'
'			
'			If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Exist(5) Then
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Click
'			End If
'
'			Browser("index:=0").Page("index:=0").Sync
'
'			If sLink<>"" Then				
'			Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("text:="&sLink,"html tag:=A").WebClick
'			Wait 5
'			End If
'
'			CaptureSnapshot()
'
'			If sCheck<>"" Then
''				If sItem="A bar" or sItem="O bar" or sItem = "Credit Alert A Bar" or sItem="Fraud" or sItem="Stolen Bar" Then
''					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
''				elseif sItem="Static IP" or sItem="Blank White Triple Format SIM" or sItem="BES 10 Gold with 2GB data" or sItem="M-Pay Bar" or sItem="Blackberry Internet Service" or sItem="BlackBerry service" or sItem="Extra 1GB data \(monthly\)"  Then
''					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//td//*[text()='"&sItem&"']/ancestor::tr[1]").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
''				Else
''					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
''				End If	
'				If sItem="A bar" or sItem="O bar" or sItem = "Credit Alert A Bar" or sItem="Fraud" or sItem="Stolen Bar" Then 
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
''				elseif sItem="Static IP" or sItem="Blank White Triple Format SIM" or sItem="BES 10 Gold with 2GB data" or sItem="M-Pay Bar" or sItem="Blackberry Internet Service" or sItem="BlackBerry service" or sItem="Extra 1GB data \(monthly\)"  Then
''					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//td//*[text()='"&sItem&"']/ancestor::tr[1]").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'				Else
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'				End If
'				AddVerificationInfoToResult  "Info" ,"Action :" & sCheck & " is done on product " &sItem
'			End If
'
'			If sCatalog <> "" Then
'				If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("There is an error with").Exist(3) Then
'					AddVerificationInfoToResult "Info" , "Pop up message is displayed successfully after selecting " & sItem & " for Prepaid Account."
'					CaptureSnapshot()
'					Exit Function
'				Else
'					AddVerificationInfoToResult "Error" , "Pop up message is not displayed successfully."
'					iModuleFail = 1
'					Exit Function
'				End If
'			End If
'
'			If sCustomise="Y" Then
'				If sLink="Bars" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebTable("column names:=.*Customise").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
'					Wait 5
'				else
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
'				End If
'				
'				'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sCustomiseItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sCustomiseValue
'			End If			
'
'			If sList<>"" Then
'				If  sList="Any"Then
'					sList="#1"
'				End If
'				If sItem="Item" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sList&"']/..").WebSelect sList
'				ElseIf sItem ="AreaCode" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebList("all items:=.*"&sList&".*").WebSelect sList
'				Else
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//*[text()='"&sItem&"']/../..//select/option[text()='"&sList&"']/..").WebSelect sList			
'				End If
'			End If
'
'			If sAddItem<>"" Then
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").WebEdit("html tag:=INPUT").Set sAddItem
'
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").Link("text:=Add Item").WebClick
''				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT","index:=0").Set sAddItem
''				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Link("text:=Add Item").WebClick
'			End If	
'
'
'If  sTextValue<>""Then
'	If sItem ="MSISDN" Then
'				sTextValue = DictionaryTest_G.Item("ACCNTMSISDN")
''                sTextValue = "447387945790"
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
'	Else
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
'
'			End If
'
'End If
'			
'			adoData.movenext
'		
'	Loop
'
'	CaptureSnapshot()
'
'	If Err.Number <> 0 Then
'            iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If
'
'
'End Function
'



'#################################################################################################
' 	Function Name : Econfig_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function Econfig_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [Econfig$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Econfig.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow
		Do while Not adoData.Eof
			sLink = adoData( "Link")  & ""
			sItem = adoData( "Item")  & ""
			sCheck = adoData( "Check")  & ""
			sList = adoData( "List")  & ""
			sCustomise= adoData( "Customise")  & ""
			'sCustomiseItem= adoData( "CustomiseItem")  & ""
			sTextValue= adoData( "TextValue")  & ""
			sAddItem= adoData( "AddItem")  & ""
			sAction = adoData( "Action")  & ""
			sCatalog = adoData( "Catalog")  & ""

			Browser("index:=0").Page("index:=0").Sync

		    If  Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Exist(3) Then
				 Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Click
			End If

			
			If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Exist(5) Then
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Click
			End If

			Browser("index:=0").Page("index:=0").Sync

			If sLink<>"" Then
				If Instr(sLink,"|") > 0 Then
					sLinkSplit = Split(sLink,"|")
				End If
				If Instr(sLink,"|") > 0 Then
					For i = 0 to Ubound(sLinkSplit)
						If sLinkSplit(i) = "RootProduct0" Then
							If DictionaryTest_G.Exists(sLinkSplit(i)) Then
									sLinkSplit(i) = Replace(sLinkSplit(i),sLinkSplit(i),DictionaryTest_G(sLinkSplit(i)))	
								elseif DictionaryTempTest_G.Exists(sLinkSplit(i)) Then
									sLinkSplit(i) = Replace(sLinkSplit(i),sLinkSplit(i),DictionaryTempTest_G(sLinkSplit(i)))
								End If
						End If
	
								If  sLinkSplit(i) = "Mobile phone service" Then
										sLinkSplit(i) = Replace("Mobile phone service","Mobile phone service","Mobile Phone Services")	
								End If
		
								If  sLinkSplit(i) = "Mobile broadband service" Then
										sLinkSplit(i) = Replace("Mobile broadband service","Mobile broadband service","Mobile broadband Services")	
								End If
	
						Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("text:="&sLinkSplit(i),"html tag:=A").WebClick
				
							If sCheck<>"" Then
								If sItem="A bar" or sItem="O bar" or sItem = "Credit Alert A Bar" or sItem="Fraud" or sItem="Stolen Bar" or sItem="20% off all international callls" Then
									sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").Exist(2)
									If sExist = "True" Then
										Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
									Else
										AddVerificationInfoToResult "Error" , "Product : " & sItem & " is not present in " & sLinkSplit(i) & " tab."
										iModuleFail = 1
										Exit Function
									End If
								Else
'									sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("class:=eCfgSpan.*","innertext:="&sItem).Exist(2)
'									If sExist = "True" Then
'										AddVerificationInfoToResult "Info" , "Product : " & sItem & " is present in " & sLinkSplit(i) & " tab."
										If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").Exist(2) Then
											
											Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
											Exit For
'										elseif Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").exist(2) Then
'											Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'											Exit For
										 elseif Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sItem&"']/..").exist(2) Then
											  Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sItem&"']/..").WebSelect sItem
												If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sItem&"']/../../..").WebEdit("html tag:=INPUT").exist(2) Then
													Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sItem&"']/../../..").WebEdit("html tag:=INPUT").Set "1"
	
													Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sItem&"']/../../..").Link("text:=Add Item").WebClick
												End If
												Exit For
'											End IF
									End If
								End If
							End If
					Next
				Else

						If sLink = "RootProduct0" Then
							If DictionaryTest_G.Exists(sLink) Then
									sLink = Replace(sLink,sLink,DictionaryTest_G(sLink))	
								elseif DictionaryTempTest_G.Exists(sLink) Then
									sLink = Replace(sLink,sLink,DictionaryTempTest_G(sLink))
								End If
						End If

								
		
								If  sLink = "Mobile phone service" Then
										sLink = Replace("Mobile phone service","Mobile phone service","Mobile Phone Services")
								End If
		
								If  sLink = "Mobile broadband service" Then
										sLink = Replace("Mobile broadband service","Mobile broadband service","Mobile broadband Services")
								End If

					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("text:="&sLink,"html tag:=A").WebClick

						If sCheck<>"" Then
							If sItem="A bar" or sItem="O bar" or sItem = "Credit Alert A Bar" or sItem="Fraud" or sItem="Stolen Bar" Then
								sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").Exist(2)
								If sExist = "True" Then
									Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
								Else
									AddVerificationInfoToResult "Error" , "Product : " & sItem & " is not present in " & sLink & " tab."
									iModuleFail = 1
									Exit Function
								End If
							Else
'								sExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("class:=eCfgSpan.*","innertext:="&sItem).Exist(2)
'								If sExist = "True" Then
'									AddVerificationInfoToResult "Info" , "Product : " & sItem & " is present in " & sLink & " tab."
									If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").Exist(2) Then
									
										Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
									
'									elseif Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").exist(2) Then
'										Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'										
									 elseif Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sItem&"']/..").exist(2) Then
										  Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sItem&"']/..").WebSelect sItem
											If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sItem&"']/../../..").WebEdit("html tag:=INPUT").exist(2) Then
												Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sItem&"']/../../..").WebEdit("html tag:=INPUT").Set "1"

												Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sItem&"']/../../..").Link("text:=Add Item").WebClick
											End If

											 Else if(sItem="Extra 1GB data \(monthly\)") Then
											
											Browser("Siebel Call Center_2").Page("Siebel Call Center_4").Frame("CfgMainFrame Frame").WebCheckBox("WebCheckBox").WebSet sCheck 

'										End If
								
									Else
										AddVerificationInfoToResult "Error" , "Product : " & sItem & " is not present in " & sLink & " tab."
										iModuleFail = 1
										Exit Function
									End If
							End If	
						End If
				End If
			End If
			End If


			If sCatalog <> "" Then
				If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("There is a conflict with").Exist(3) Then
					AddVerificationInfoToResult "Info" , "Pop up message is displayed successfully after selecting " & sItem & " for Prepaid Account."
					CaptureSnapshot()
					Exit Function
				Else
					AddVerificationInfoToResult "Error" , "Pop up message is not displayed successfully."
					iModuleFail = 1
					Exit Function
				End If
			End If

			If sCustomise="Y" Then
				If sLink="Bars" Then
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebTable("column names:=.*Customise").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
				else
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
				End If
				
				'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sCustomiseItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sCustomiseValue
			End If			

			If sList<>"" Then
				If  sList="Any"Then
					sList="#1"
				End If
				If sItem="Item" Then
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sList&"']/..").WebSelect sList
				ElseIf sItem ="AreaCode" Then
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebList("all items:=.*"&sList&".*").WebSelect sList
				Else
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//*[text()='"&sItem&"']/../..//select/option[text()='"&sList&"']/..").WebSelect sList			
				End If
			End If

			If sAddItem<>"" Then
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").WebEdit("html tag:=INPUT").Set sAddItem

				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").Link("text:=Add Item").WebClick
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT","index:=0").Set sAddItem
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Link("text:=Add Item").WebClick
			End If	



'			If sTextValue<>"" Then
'				sTextValue = DictionaryTest_G.Item("ACCNTMSISDN")
''                sTextValue = "447387945790"
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
'			End If

'*******Don not delete this code
			If  sTextValue<>""Then
				If sItem ="MSISDN" Then
							sTextValue = DictionaryTest_G.Item("ACCNTMSISDN")
			'                sTextValue = "447387945790"
							Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
				Else
							Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
			
						End If

			End If
'*******Don not delete this code


			adoData.movenext
		
	Loop

	CaptureSnapshot()

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : EconfigVerify_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function EconfigVerify_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [EconfigVerify$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Econfig.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow
		Do while Not adoData.Eof
			sLink = adoData( "Link")  & ""
			sItem = Trim(adoData( "Item")  & "")
			sExist = adoData( "Exist")  & ""
			If sExist="Y" Then
				sExist=True
			else
				sExist=False
			End If
			sAction = adoData( "Action")  & "" 


			Browser("index:=0").Page("index:=0").Sync

		    If  Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Exist(3) Then
				 Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Click
			End If

			
			If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Exist(5) Then
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Click
			End If

			Browser("index:=0").Page("index:=0").Sync

			If sLink<>"" Then				
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("text:="&sLink,"html tag:=A").WebClick
			End If
			Browser("index:=0").Page("index:=0").Sync

			If sExist <> "" Then
'					If sExist = True Then
						actExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("class:=eCfgSpanAvailable","innertext:="&sItem).Exist(2)
'					else if sExist = False Then
'						actExist = Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("class:=eCfgSpanExcluded","innertext:="&sItem).Exist
'					End If
			End If

'			If cstr(actExist)=cstr(True) Then
'				actExist1="Y"
'			elseif  cstr(actExist)=cstr(False) Then
'				actExist1="N"
'			End If

			If cstr(sExist) = cstr(actExist) Then
				AddVerificationInfoToResult  "Pass" , "Existence of " & sItem & " is " & sExist & " as expected"
			Else
				AddVerificationInfoToResult  "Fail" , "Existence of " & sItem & " is not as expected."
				iModuleFail = 1
			End If
			CaptureSnapshot()


			If sAction="Done"  Then

				 SiebApplication("Siebel Call Center").SetTimeOut(10)

				On error resume next
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Done").Click

				If Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Exist(2) Then
					Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Click
				End If

				Browser("index:=0").Page("index:=0").Sync

			End If


			Err.Clear
	


			adoData.movenext
		
	Loop

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function



'#################################################################################################
' 	Function Name : EconfigDoneButton_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function EconfigDoneButton_fn()

	call SetObjRepository ("Orders",sProductDir & "Resources\")

				 SiebApplication("Siebel Call Center").SetTimeOut(10)
				Reporter.Filter = 3
				On error resume next
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Done").Click

				If Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Exist(2) Then
					Browser("Siebel Call Center_2").Dialog("Message from webpage").WinButton("OK").Click
				End If

				Browser("index:=0").Page("index:=0").Sync


			 Reporter.Filter = 0
			Err.Clear
	
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : OrderFormEntry1_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderFormEntry1_fn()

	Dim adoData	 

    strSQL = "SELECT * FROM [OrderFormEntry$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	sDisconnectionReason = adoData( "DisconnectionReason")  & ""
	sReturnReason =  adoData( "ReturnReason")  & ""
	sJourney = adoData( "Journey")  & ""
	sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If


	'Flow
	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************


	If  sJourney = "Buyback" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Buyback").SiebClick False
		Exit Function
	End If


		' Code to capture errror message while making 1 On Account Payments.
	If Browser("Siebel Call Center").Dialog("Siebel").Static("The maximum amount for").Exist(3) Then
		CaptureSnapshot()
		strErrorMessage = Browser("Siebel Call Center").Dialog("Siebel").Static("The maximum amount for").GetROProperty("text")
		If Instr(strErrorMessage,"maximum amount for the On Account payment") > 0 Then
		   AddVerificationInfoToResult  "Error Message" , "Error Message:  " & strErrorMessage & "is displayed correctly after making 2 On Account Payments."
		Else
			AddVerificationInfoToResult  "Error Message" , "Error Message is not displayed successfully after making 2 On Account Payments."
		End If
		Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
		Exit Function
	End If
	
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCheckbox("Terms and conditions").SiebSetCheckbox "ON"

    If sDisconnectionReason <> "" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Disconnection reason").Exist(2)=False Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click
		End If
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Disconnection reason").SiebSelect sDisconnectionReason
	End If

	If sReturnReason <> "" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Return reason").Exist(2)=False Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click
		End If
		 SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Return reason").SiebSelect sReturnReason
	  End If
	
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Verify").SiebClick sPopUp

'	Browser("index:=0").Page("index:=0").Sync
  'SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Submit").SiebClick sPopUp

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : OrderFormEntry_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderFormEntry_fn()

	'sOR
Dim adoData	 

    strSQL = "SELECT * FROM [OrderFormEntry$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	sDisconnectionReason = adoData( "DisconnectionReason")  & ""
	sReturnReason =  adoData( "ReturnReason")  & ""
	sJourney = adoData( "Journey")  & ""
	sClickSubmit =  adoData( "ClickSubmit")  & ""
	sPopUp = adoData( "Popup")  & ""	

		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If


	'Flow
	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If


	If  sJourney = "Buyback" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Buyback").SiebClick sPopUp
		Exit Function
	End If


		' Code to capture errror message while making 1 On Account Payments.
       If Browser("Siebel Call Center").Dialog("Siebel").Static("The maximum amount for").Exist(3) Then
		    CaptureSnapshot()
		   strErrorMessage = Browser("Siebel Call Center").Dialog("Siebel").Static("The maximum amount for").GetROProperty("text")
		   If Instr(strErrorMessage,"maximum amount for the On Account payment") > 0 Then
			   AddVerificationInfoToResult  "Error Message" , "Error Message:  " & strErrorMessage & "is displayed correctly after making 2 On Account Payments."
			Else
			AddVerificationInfoToResult  "Error Message" , "Error Message is not displayed successfully after making 2 On Account Payments."
		   End If
			Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
			Exit Function
	 End If
	
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCheckbox("Terms and conditions").SiebSetCheckbox "ON"

    If sDisconnectionReason <> "" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Disconnection reason").Exist(2)=False Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click
		End If
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Disconnection reason").SiebSelect sDisconnectionReason
	End If


	  If sReturnReason <> "" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Return reason").Exist(2)=False Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click
		End If
		 SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Return reason").SiebSelect sReturnReason
	  End If


	
'	SiebApplication("Siebel Call Center").SetTimeOut(10)
'    Reporter.Filter = 3
'	On error resume next
	If Ucase(sJourney) = "VALIDATION" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Verify").SiebClick False
	Else
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Verify").SiebClick sPopUp
	End If
'	
'    Reporter.Filter = 0
'    Err.Clear


	Browser("index:=0").Page("index:=0").Sync
	
		If sClickSubmit <>"N" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Submit").SiebClick sPopUp
		End If


		If sJourney = "CreditAccount" Then
				sStatus=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Status").SiebGetRoProperty("activeitem")
				i = 0
		
				Do while sStatus <> "Complete" and i <=50
			
					wait 5
			
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebMenu("Menu").Select "ExecuteQuery"
			
					sStatus=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Status").SiebGetRoProperty("activeitem")
			
					If sStatus = "Complete" Then
						Exit Do
					End If
			
					i = i+1
			
				Loop
		
				If   sStatus <> "Complete"Then
						  iModuleFail = 1
							AddVerificationInfoToResult  "Error" , "Order Sattaus is not Complete"
				End If
		
			End If

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : CancelOrder_fn
' 	Description : This function is used to cancel the order
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CancelOrder_fn()

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

    
	'Flow
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************
	sOrderStatus = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Status").GetROProperty("activeitem")
	If sOrderStatus = "Pending" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Cancel reason").SiebSelect "No Reason"
	End If

    	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Cancel Order").SiebClick False
		Browser("index:=0").Page("index:=0").Sync

	If sOrderStatus = "Open" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Cancel reason").SiebPicklist("Cancel reason").SiebSelect "No Reason"
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Cancel reason").SiebButton("Go").SiebClick False
	End If

	AddVerificationInfoToResult  "Cancel Order" , "Order is cancelled successfully."

      CaptureSnapshot()
	   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : CaptureOrderDetails_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CaptureOrderDetails_fn()

Dim adoData	  


	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	Browser("index:=0").Page("index:=0").Sync

	'Flow
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"

	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************

	sOrderNo=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Order no.").SiebGetRoProperty("text")
	If DictionaryTest_G.Exists("OrderNo") Then
		DictionaryTest_G.Item("OrderNo")=sOrderNo
	else
		DictionaryTest_G.add "OrderNo",sOrderNo
	End If

	AddVerificationInfoToResult  "Info" , "OrderNo is " & sOrderNo

	sStatus=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Status").SiebGetRoProperty("activeitem")


	If DictionaryTest_G.Exists("Status") Then
		DictionaryTest_G.Item("Status")=sStatus
	else
		DictionaryTest_G.add "Status",sStatus
	End If
	AddVerificationInfoToResult  "Info" , "Status is " & sStatus

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function



'#################################################################################################
' 	Function Name : CreditVettingResults_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CreditVettingResults_fn()
   On error resume next
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If
'   SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************

	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Credit Vetting Results").SiebMenu("Menu").Select "ExecuteQuery"

'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items").SiebApplet("Credit Vetting Results").SiebMenu("Menu").Select 
'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Credit Vetting Results").SiebCheckbox("Override credit vet").Highlight
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Credit Vetting Results").SiebCheckbox("Override credit vet").SiebSetCheckbox "ON"
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Credit Vetting Results").SiebCheckbox("Override credit vet").SiebSetCheckbox "ON"
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Credit Vetting Results").SiebMenu("Menu").Select "ExecuteQuery"
	'ighlight
	CaptureSnapshot()
   
'	If Err.Number <> 0 Then
'            iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If


End Function



'#################################################################################################
' 	Function Name : CustomiseProduct_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CustomiseProduct_fn()

	'Flow


Dim adoData	  
    strSQL = "SELECT * FROM [CustomiseProduct$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sSelectItem = adoData( "SelectItem")  & ""
			If DictionaryTest_G.Exists(sSelectItem) Then
			sSelectItem = Replace(sSelectItem,sSelectItem,DictionaryTest_G(sSelectItem))
	
	elseif DictionaryTempTest_G.Exists(sSelectItem) Then
			sSelectItem = Replace(sSelectItem,sSelectItem,DictionaryTempTest_G(sSelectItem))
		End If 
	

	Browser("index:=0").Page("index:=0").Sync
	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If	
	Browser("index:=0").Page("index:=0").Sync

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************
'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCheckbox("Terms and conditions").SiebSetCheckbox "ON"
	Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items")
	Do while Not adoData.Eof
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""
'		sLocateColValue = Replace(sLocateColValue,"RootProduct0",DictionaryTest_G.Item("ROOT_PRODUCT"))
		iIndex = adoData( "Index")  & ""
		If iIndex="" Then
			iIndex=0
'		else
'			iIndex=Cint(iIndex)
		End If

		If sLocateCol <> "" Then
			res=LocateColumns (objApplet,sLocateCol,sLocateColValue,iIndex)
		End If
		adoData.MoveNext
	Loop	

'	If sSelectItem="InstalledId" Then
'		MSISDN = DictionaryTest_G.Item("ACCNTMSISDN")
'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectInstalledIdRow MSISDN
'	else
'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectProductRow sSelectItem,0
'	End If	
	
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebButton("Customise").SiebClick False

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function




'#################################################################################################
' 	Function Name : AddProduct_fn
' 	Description : 
'   Created By :  Ankit/Tarun
'	Creation Date :        
'##################################################################################################
Public Function AddProduct_fn()


	'Flow


	Dim adoData	  
    strSQL = "SELECT * FROM [AddProduct$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sAddProduct = adoData( "AddProduct")  & ""
	sQuantity = adoData( "Quantity")  & ""
	sProductName = adoData( "ProductName")  & ""
	sMultiHandset = adoData( "MultiHandset")  & ""
	sProductID =  adoData( "ProductID")  & ""
	Browser("index:=0").Page("index:=0").Sync

	Set obj=Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("innertext:="&sAddProduct&".*","html tag:=TR","index:=0")

	' If obj.Exist(10) Then
		' obj.Link("innertext:=Add","index:=0").Click
	' Else
		' AddVerificationInfoToResult  "Fail" ,"Product: " & sAddProduct & " is not available on page."
		' iModuleFail = 1
		' Exit Function
	' End If
	
	If obj.Exist(7) Then
		obj.Link("innertext:=Add").WebClick
	Else
		AddVerificationInfoToResult  "Fail" ,"Product: " & sAddProduct & " is not available on page."
		iModuleFail = 1
		Exit Function
	End If

	If sProductName <>""Then
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebList("PopupQueryCombobox").WebSelect "Product name"
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebEdit("PopupQuerySrchspec").WebSet sProductName
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Go").Click
	End If

		If sProductID <> "" Then
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebList("PopupQueryCombobox").WebSelect "Product Id" 
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebEdit("PopupQuerySrchspec").WebSet sProductID
		Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Go").Click
	End If


		If sMultiHandset = "Y" Then

				Browser("index:=0").Page("index:=0").Sync
				Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebEdit("Quantity.0").WebSet sQuantity
				Browser("index:=0").Page("index:=0").Sync
				Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow_3").Frame("_swepopcontent Frame").WebEdit("Quantity.1").WebSet sQuantity
				CaptureSnapshot()
				Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Add").Click
				AddVerificationInfoToResult  "Pass" ,"Product Qty Added: "
				Browser("index:=0").Page("index:=0").Sync
				Exit Function
			End If

						' If (Ucase(sEnv) = "SUP02" OR Ucase(sEnv) = "E7")  Then ''''Only for SUP02
											If  (instr(sAddProduct,"Insurance") >0 Or instr(sAddProduct,"Vodafone") >0 Or instr(sAddProduct,"Fixed Service") >0 Or instr(sAddProduct,"Blank White Triple Format SIM") >0) Then
													Exit Function
											Else
												Browser("index:=0").Page("index:=0").Sync
													Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebEdit("Quantity.0").WebSet sQuantity
													Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Add").Click
													Browser("index:=0").Page("index:=0").Sync
													AddVerificationInfoToResult  "Pass" ,"Product Qty Added: "
													CaptureSnapshot()
													Exit Function
										
											End If		
									'End If
									''''For other ENvs
										Browser("index:=0").Page("index:=0").Sync
										Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").WebEdit("Quantity.0").WebSet sQuantity
										Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("Add").Click
										Browser("index:=0").Page("index:=0").Sync
										AddVerificationInfoToResult  "Pass" ,"Product Qty Added: "
										CaptureSnapshot() 
					

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : Buyback_fn
' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Buyback_fn()

	'sOR
	
	call SetObjRepository ("Account",sProductDir & "Resources\")

	'Flow
	Browser("index:=0").Page("index:=0").Sync
	If  Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(5) Then
		Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Click
	End If

	If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist(2) Then
		Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
	End If

	If  Browser("Retail Phone Shop").Page("Retail Phone Shop").WebElement("Buyback").Exist(10) Then
		AddVerificationInfoToResult  "Pass" , "Page is displayed successfully after clicking on Buyback button."
	Else
		AddVerificationInfoToResult  "Fail" , "Page is not displayed successfully after clicking on Buyback button."
		iModuleFail = 1
	End If

	Browser("Retail Phone Shop").highlight

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : OrderBillingProfileSelection_fn
' 	Description :  This function is used to select billing profile while doing Pre/Post migration in Orders page
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderBillingProfileSelection_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [BillingProfileSelection$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sBillingProfileSelection = adoData( "BillingProfile")  & ""
	
	Browser("index:=0").Page("index:=0").Sync

	

	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing profile").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing profile").OpenPopup
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile").SiebPicklist("SiebPicklist").SiebSelect "Postpay/ Prepay"
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile").SiebText("SiebText").SiebSetText sBillingProfileSelection
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile").SiebButton("Go").SiebClick False
		AddVerificationInfoToResult  "Info" , "Billing Profile for " & sBillingProfileSelection & " is selected successfully."
	Else
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").click
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing profile").OpenPopup
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile").SiebPicklist("SiebPicklist").SiebSelect "Postpay/ Prepay"
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile").SiebText("SiebText").SiebSetText sBillingProfileSelection
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile").SiebButton("Go").SiebClick False
		AddVerificationInfoToResult  "Info" , "Billing Profile for " & sBillingProfileSelection & " is selected successfully."
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : OrderNumberSelection_fn
' 	Description :  This function is used to select the order number from orders tab, check the order status and click on customer account button.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderNumberSelection_fn()


	Dim adoData	  
'    strSQL = "SELECT * FROM [CaptureDataOutputSheet$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\OutputSheet.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sOrderNumber =  DictionaryTest_G.Item("OrderNo")
'  sOrderNumber = "SBL-1000000000292744"
	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Sales Order Screen" ' clicking on Orders tab
	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - All Orders View (Sales)","L2"

	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebButton("Query").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebList("List").SiebText("Order no.").SiebSetText sOrderNumber
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebButton("Go").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebList("List").SiebDrillDownColumn "Order no.",0
	
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3" ' cliking on line items
		Browser("index:=0").Page("index:=0").Sync

    		sOrderStatus = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Status").SiebGetRoProperty("activeitem") ' Order status
	
		If sOrderStatus = "Complete" Then
			AddVerificationInfoToResult  "Info" ,"Order status is completed"
		Else
			AddVerificationInfoToResult  "Fail" ,"Order status is : " & sOrderStatus
			 iModuleFail = 1
		End If
	Else
		AddVerificationInfoToResult  "Fail" ,"Orders Page is not displayed successfully."
		 iModuleFail = 1
		 Exit Function
	End If
	CaptureSnapshot()

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Customer Account").SiebClick False ' clicking on CustonerAccount button.
   
	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : PaymentDetails_fn
' 	Description :  This function is used to select the order number from orders tab, check the order status and click on customer account button.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function PaymentDetails_fn()

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Payment Detail - Balance").SiebText("Mobile no.").OpenPopUp
	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Pick Asset").SiebButton("OK").Click
	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Detail - Balance").SiebButton("Reserve Funds").SiebClick False


	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : CostOverride_fn
' 	Description : This function performs cost override for Office 365 products
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function CostOverride_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [CostOverride$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	sCostOverrideValue = adoData( "CostOverrideValue")  & ""

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow
' sMonthlyCost =SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCurrency("Monthly cost").GetROProperty("text")
'AddVerificationInfoToResult  "Info" ,"MonthlyCost before Cost Override is : " & sMonthlyCost

		If sCostOverrideValue = "MonthlyCost"  Then
			 sMonthlyCost =SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCurrency("Monthly cost").GetROProperty("text")
			AddVerificationInfoToResult  "Info" ,"MonthlyCost before Cost Override is : " & sMonthlyCost
			 a=len(sMonthlyCost)
		'		b=a-3
				sCostOverride = "-"&mid(sMonthlyCost,2,a)
				CaptureSnapshot()
				'sCostOverride = "-"&sMonthlyCost
		Else sCostOverride =  sCostOverrideValue
		sMonthlyCost =SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCurrency("Monthly cost").GetROProperty("text")
		AddVerificationInfoToResult  "Info" ,"MonthlyCost before Cost Override is : " & sMonthlyCost
		CaptureSnapshot()

		End If

		 sFinalRecurringCostbeforeDiscount = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebCalculator("Final Recurring Cost before discounts)").GetROProperty("text")
		AddVerificationInfoToResult  "Info" ,"FinalRecurringCostbeforeDiscount before Cost Override is : " & sFinalRecurringCostbeforeDiscount
		
		 sFinalRecurringCostIncDiscount = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebCalculator("Final Recurring Cost (inc. discounts)").GetROProperty("text")
		 AddVerificationInfoToResult  "Info" ,"FinalRecurringCostIncDiscount before Cost Override is : " & sFinalRecurringCostIncDiscount

   	'''performing Cost Override
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebCalculator("Cost override inc. VAT").SetText sCostOverride
		Browser("Siebel Call Center_2").Page("Siebel Call Center").Frame("View Frame").WebElement("Cost override inc. VAT:").Click
		CaptureSnapshot()

				sMonthlyCost =SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCurrency("Monthly cost").GetROProperty("text")
                				
				 sFinalRecurringCostbeforeDiscount = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebCalculator("Final Recurring Cost before discounts)").GetROProperty("text")
				AddVerificationInfoToResult  "Info" ,"FinalRecurringCostbeforeDiscount beafterfore Cost Override is : " & sFinalRecurringCostbeforeDiscount
				
				 sFinalRecurringCostIncDiscount = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebCalculator("Final Recurring Cost (inc. discounts)").GetROProperty("text")
				 AddVerificationInfoToResult  "Info" ,"FinalRecurringCostIncDiscount after Cost Override is : " & sFinalRecurringCostIncDiscount

					If sCostOverrideValue = "MonthlyCost" and sMonthlyCost = "£0.00" Then
						AddVerificationInfoToResult  "Info" ,"Monthly Cost after Cost Override is : " & sMonthlyCost
					ElseIf sCostOverrideValue <> "MonthlyCost" Then
						sReducedPrice = sMonthlyCost - sCostOverrideValue
						AddVerificationInfoToResult  "Info" ,"Reduced Monthly Cost after Cost Override is : " & sMonthlyCost
					Else
					AddVerificationInfoToResult  "Fail" ,"Monthly Cost after Cost Override is not 0"
					 iModuleFail = 1
    				End If

End Function



'#################################################################################################
' 	Function Name : OrderDescriptionVerfy_fn
' 	Description :  This function is used to verify whteher Order Back Office Description is visible or not and also to validate it.
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################

Public Function OrderDescriptionVerify_fn()

'               'sOR
Dim adoData      

    strSQL = "SELECT * FROM [OrderDescriptionVerify$] WHERE  [RowNo]=" & iRowNo
                Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

                'sOR
                call SetObjRepository ("Orders",sProductDir & "Resources\")

                sExpectedOrderDescription = adoData( "ExpectedOrderDescription")  & ""             
                sAboutRecord = adoData("AboutRecord")  & ""
				sServiceAccountValidation =  adoData("ServiceAccountValidation")  & ""
				sLastNameValidation =   adoData("LastNameValidation")  & ""

                'Flow
                activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
                If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
                                SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
                End If

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************
'                
                If sAboutRecord = "Y" Then
                             SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebMenu("Menu").SiebSelectMenu "About Record"
							CaptureSnapshot
                End If    



				If  sExpectedOrderDescription<>""Then			
								If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebTextArea("Back office description").Exist(2) = False Then
											SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click
								End If
					
							'Captures Back Office Description text
								sOrderDescription =        SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebTextArea("Back office description").GetROProperty("text")
																																																													
							  If (Ucase(sOrderDescription) = Ucase(sExpectedOrderDescription)) Then
								  AddVerificationInfoToResult  "Pass" , " Existence of text : " & sOrderDescription & " as expected"
								  CaptureSnapshot()
							   Exit Function
							else
								  iModuleFail = 1
									AddVerificationInfoToResult  "Error" , "Existence of text : " & sOrderDescription & " but expected was " & sEnable
									CaptureSnapshot()
								 Exit Function
						 End If
            	End If    

				If  sServiceAccountValidation = "NotReadOnly" Then
					If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").exist(2) =False Then
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click						
					End If
					strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").getRoProperty("isenabled")
					CaptureSnapshot
						If  strEnable = True Then
								AddVerificationInfoToResult "Pass","Service Account is editable.."
							Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Service Account is not editable."
								Exit Function

						End If

						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").OpenPopup
						If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").exist(2)Then
							rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebList("List").SiebListRowCount
							CaptureSnapshot

									If rowCnt <> 0 Then
										AddVerificationInfoToResult "Pass","Rows are displayed. Service Account is editable.."
									Else
										iModuleFail = 1
										AddVerificationInfoToResult "Fail","No rows displayed. Service Account is not editable."
										Exit Function
									End If
								SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Account").SiebButton("OK").Click				
						End If
				End If	
						
				If  sLastNameValidation = "NotReadOnly" Then
					strEnable = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Last name").getROProperty("isenabled")
					CaptureSnapshot
					If  strEnable = True Then
								AddVerificationInfoToResult "Pass","Last Name is editable.."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Last Name is not editable."
						Exit Function
						End If
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Last name").OpenPopup
						If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Contact").exist(2)Then
							rowCnt = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Contact").SiebList("List").SiebListRowCount
							CaptureSnapshot

									If rowCnt <> 0 Then
										AddVerificationInfoToResult "Pass","Rows are displayed. Last Name is editable.."
									Else
										iModuleFail = 1
										AddVerificationInfoToResult "Fail","No rows displayed. Last Name is not editable."
										Exit Function
									End If
								SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Contact").SiebButton("OK").Click				
						End If
				End If

				If  sServiceAccountValidation = "ReadOnly" Then

						If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").exist(2) =False Then
							SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click						
						End If
						strEnable =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Service account").getRoProperty("isenabled")
						CaptureSnapshot
						If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Service Account is not editable.."
							Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Service Account is  editable."
								Exit Function
						End If

				End If

				If  sLastNameValidation = "ReadOnly" Then
					strEnable = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Last name").getROProperty("isenabled")
					CaptureSnapshot
					If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Last Name is not editable.."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Last Name is editable."
						Exit Function
					End If
				End If

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If			

	End Function

'#################################################################################################
' 	Function Name : OrderNumberSelectionBulkModify_fn
' 	Description :  This function is used to select the order number from orders tab, check the order status and click on customer account button.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderNumberSelectionBulkModify_fn()


	Dim adoData	  

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sOrderNumber =  DictionaryTest_G.Item("OrderNo")

	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Sales Order Screen" ' clicking on Orders tab

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - All Orders View (Sales)","L2"

	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebButton("Query").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebList("List").SiebText("Order no.").SiebSetText sOrderNumber
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebButton("Go").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebList("List").SiebDrillDownColumn "Order no.",0
	End If

	CaptureSnapshot

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name :  OrderVerify_fn
' 	Description : This function takes screenshot of popup msg
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function OrderVerify_fn()

	'sOR
Dim adoData	 

'    strSQL = "SELECT * FROM [OrderFormEntry$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")


	'Flow
'	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
'	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
'	End If

	
'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Verify").Click

'If  Browser("Siebel Call Center").Dialog("Message from webpage").Exist(2) Then
'		Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click



			If  Browser("Siebel Call Center").Dialog("Siebel").Static("The MSISDN you have entered").Exist(5)Then
				AddVerificationInfoToResult  "Pass" , " Existence of popup as expected"
				CaptureSnapshot()
			
				Browser("Siebel Call Center").Dialog("Siebel").WinButton("OK").Click
		
				If Browser("Siebel Call Center").Dialog("Message from webpage").Exist(2) Then
					Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
				End If
			else
				iModuleFail = 1
					strPopup =Browser("Siebel Call Center").Dialog("Siebel").Static("The MSISDN you have entered").getROProperty("text")
					AddVerificationInfoToResult  "Error" , "Existence of text : The MSISDN you have entered: but expected was " & strPopup
					CaptureSnapshot()
			End if
'End If


	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : OrderReportVerify_fn
' 	Description :  This function is used to verify the Report (PDF) which gets generated for new Provision
'   Created By :  Mohit
'	Creation Date :        
'##################################################################################################
Public Function OrderReportVerify_fn()

	Dim adoData	  

	Call SetObjRepository ("Orders",sProductDir & "Resources\")

'	Browser("creationtime:=0").Refresh
	Wait 2
	sAccount_Type = DictionaryTest_G.Item ("sAccount_Type")

	If sEnv = "E7" Then
		MenuSelection = "Reports\\Generate Contract - Mobile"
	Else
		If Instr (Ucase(sAccount_Type), Ucase("Consumer")) Then
			MenuSelection = "Reports\\Generate Contract - Consumer Mobile"
		Else
			MenuSelection =  "Reports\\Generate Contract - Enterprise Mobile"
		End If
	End If

	SiebApplication("Siebel Call Center").SiebMenu("Reports").Select MenuSelection
	Wait 4


	SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Output Type").SiebButton("Submit").SiebClick False

'	Wait 4
	If  Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(1) Then
		Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Click
	End If
'	Wait 3
	If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist(0) Then
		Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
	End If
'	Wait 5
	If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist(0) Then
		Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
	End If
'	Wait 2

	If Browser("Certificate Error: Navigation").WinButton("To help protect your security,").Exist(2) Then
'		print "security box exist"
		Set obj=createobject("mercury.devicereplay")
		Set  winBtn = Browser("Certificate Error: Navigation").WinButton("To help protect your security,")
        getX = winBtn.GetROProperty("abs_x")
		getY = winBtn.GetROProperty("abs_y")

		obj.Highlight	
		obj.MouseClick getX+5, getY+5,RIGHT_MOUSE_BUTTON
'        Browser("Certificate Error: Navigation_2").WinButton("To help protect your security,").Click getX+5,getY+5,micRightBtn
'		Browser("Certificate Error: Navigation_2").WinMenu("ContextMenu").Select "Download File..."
		Browser("Certificate Error: Navigation").WinMenu("ContextMenu").Select "Download File..."
	End If


	If Browser("Siebel Call Center_Report").Page("Siebel Call Center").WebElement("spacer").Exist(5) Then
		Browser("Siebel Call Center_Report").Close
		Wait 2
		Browser("SiebWebPopupWindow").CloseAllTabs
		Wait 2

		SiebApplication("Siebel Call Center").SiebMenu("Reports").Select MenuSelection

		SiebApplication("Siebel Call Center").SiebScreen("SiebScreen").SiebView("SiebView").SiebApplet("Output Type").SiebButton("Submit").SiebClick False

		If  Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(1) Then
			Browser("Certificate Error: Navigation").Page("Certificate Error: Navigation").Link("Continue to this website").Click
		End If
	'	Wait 3
		If Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").Exist(0) Then
			Browser("micclass:=Browser","name:=Certificate Error: Navigation Blocked").Dialog("micclass:=Dialog","text:=Security Warning").WinButton("micclass:=WinButton","text:=Yes").Click
		End If
	'	Wait 5



	End If

'	Set WshShell = CreateObject("WScript.Shell")
'	For i = 0 to 3
'		WshShell.SendKeys "{left}"
'		Wait 1
'	Next
'	WshShell.SendKeys "~"

	Browser("Certificate Error: Navigation").Dialog("File Download").Activate
	Browser("Certificate Error: Navigation").Dialog("File Download").WinButton("Open").Highlight	
	Browser("Certificate Error: Navigation").Dialog("File Download").WinButton("Open").Click
	Wait 3
	Window("Adobe Acrobat Reader DC").WinObject("object class:=AVL_AVView", "text:=AVPageView").Type micCtrlDwn + "a" + micCtrlUp
	Wait 4
	Window("Adobe Acrobat Reader DC").WinObject("object class:=AVL_AVView", "text:=AVPageView").Type micCtrlDwn + "c" + micCtrlUp
	Wait 4
	
	Set clip = CreateObject("Mercury.Clipboard" )
	strText = clip.GetText
	clip.Clear

	AccountNo = DictionaryTest_G.Item ("AccountNo")
	MSISDN = DictionaryTest_G.Item ("MSISDN")
'	sAccount_Type = DictionaryTest_G.Item ("sAccount_Type")

	If Instr(strText, AccountNo) Then
		AddVerificationInfoToResult "Info", "Account #" & AccountNo &"verified successfully"
		If Instr(strText, MSISDN) Then
			AddVerificationInfoToResult "Info", "MSISDN #" & MSISDN & "verified successfully"
			If Instr(strText, sAccount_Type) Then
				AddVerificationInfoToResult "Info", "Account Type " & sAccount_Type & " verified successfully"
			Else
				AddVerificationInfoToResult "Info", "Account Type " & sAccount_Type & " not verified successfully."
				iModuleFail = 1
			End If
		Else
				AddVerificationInfoToResult "Info", "MSISDN #" & MSISDN & " not verified successfully."
				iModuleFail = 1
		End If
	Else
				AddVerificationInfoToResult "Info", "Account # " & AccountNo & " not verified successfully."
				iModuleFail = 1
	End If
	
	Wait 2
'	Browser("Certificate Error: Navigation").Close
	

	If Err.Number <> 0 Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function
'#################################################################################################
' 	Function Name : OneNetBusinessCreditVet_fn
' 	Description :  This function is used to verify the Report (PDF) which gets generated for new Provision
'   Created By :  Mohit
'	Creation Date :        
'##################################################################################################
Public Function OneNetBusinessCreditVet_fn()

	Dim adoData	  
    strSQL = "SELECT * FROM [CreditVetClick$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	sPopUp = adoData( "Popup")  & ""	
	If Ucase(sPopUp)="NO" Then
		sPopUp="FALSE"
		sPopUp=Cbool(sPopUp)
	End If

'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	Browser("index:=0").Page("index:=0").Sync

    '*****************Web Element********************
Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
Browser("index:=0").Page("index:=0").Sync
'*****************Web Element********************
    
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Credit Vet").SiebClick sPopUp

'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Catalogue").SiebApplet("Orders").SiebButton("Credit Vet").SiebClick sPopUp
	Browser("index:=0").Page("index:=0").Sync
    
End Function
'#################################################################################################
' 	Function Name : RetrieveBBIP
' 	Description :  This function is used to verify the Report (PDF) which gets generated for new Provision
'   Created By :  Mohit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveBBIP()

	Dim sDate : sDate = Day(Now)
	Dim sMonth : sMonth = Month(Now)
	Dim sYear : sYear = Year(Now)
	Dim sHour : sHour = Hour(Now)
	Dim sMinute : sMinute = Minute(Now)
	Dim sSecond : sSecond = Second(Now)

	fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sHour & sMinute & sSecond)
	BBIPNum = "BBIP" & fnRandomNumberWithDateTimeStamp
	
	RetrieveBBIP = BBIPNum
'    DictionaryTest_G.Add "BBIPNum", BBIPNum
'	msgbox BBIPNum
End Function


'#################################################################################################
' 	Function Name :  CopyOrderLineItems_fn
' 	Description : This function  is used to click on Copy button on oreder line items page and enter the quantity in bulk copy window.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CopyOrderLineItems_fn()

	'sOR
Dim adoData	 

'    strSQL = "SELECT * FROM [OrderFormEntry$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebButton("Copy").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebButton("Copy").SiebClick False

			Browser("index:=0").Page("index:=0").Sync

			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Bulk Copy").Exist(10) Then
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Bulk Copy").SiebCalculator("Quantity").SetText "2"
				CaptureSnapshot()
				AddVerificationInfoToResult  "Info" ,"Quantity is added succesfully in Bulk Copy window."
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Bulk Copy").SiebButton("Go").SiebClick False
			Else
				AddVerificationInfoToResult  "Error" ,"Copy button is not clicked successfully on Order line items page."
				iModuleFail = 1
			End If
		Else
			AddVerificationInfoToResult  "Error" ,"Copy button is not present on Order line items page."
			iModuleFail = 1
		End If

'SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Bulk Copy").SiebCalculator("Quantity").SetText


	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function



'#################################################################################################
' 	Function Name : PromotionUpgrade_fn
' 	Description : This function performs Tariff Migration,Pre/Post
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################

Public Function PromotionSelect_fn()


	'Get Data
'	Dim adoData
'	Dim sStartingWith

'    strSQL = "SELECT * FROM [PromotionUpgrade$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

'SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Promotion").SiebPicklist("SiebPicklist").Select
'SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Promotion").SiebText("SiebText").SetText
'SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Promotion").SiebList("List").DrillDownColumn
''SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Promotion").SiebButton("Go").Click
		Browser("index:=0").Page("index:=0").Sync
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Promotion").Exist (5) Then
			CaptureSnapshot()
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Pick Promotion").SiebButton("OK").Click
			Browser("index:=0").Page("index:=0").Sync
			 AddVerificationInfoToResult  "Pass" , "Promotion Select Applet found"
		Else
			 iModuleFail = 1
			 AddVerificationInfoToResult  "Fail" , "Promotion Select Applet not found"
			 Exit Function
		End If
		
		
		
		
		CaptureSnapshot()

		Browser("index:=0").Page("index:=0").Sync

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


''#################################################################################################
' 	Function Name : OrdersFulfilmentTab_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function OrdersFulfilmentTab_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [OrdersFulfilment$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	sStockCheck = adoData( "StockCheck")  & ""
	sReserve = adoData( "Reserve")  & ""
    sStoreAddress = adoData("StoreAddress") & ""
	sStoreID = adoData("StoreID") & ""
	sStockCheckTextBox = adoData("StockCheckTextBox") & ""

		sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

	'Flow
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Shipment Line Detail View (Sales)","L3"
	Browser("index:=0").Page("index:=0").Sync

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Fulfillment").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************


    If sStoreAddress = "Y" Then		
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Fulfillment").SiebText("Store address").OpenPopup
				If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Store code").Exist(5)Then
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Store code").SiebText("SiebText").SetText sStoreID
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Store code").SiebButton("Go").Click
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Store code").SiebButton("OK").Click
					AddVerificationInfoToResult "Pass" , "Successfully Selected Store Address"
				Else
					AddVerificationInfoToResult "Error" , "Store Address Pop Up did not occur"
					iModuleFail = 1
					CaptureScrenshot()
					Exit Function
				End If
			End If

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Line Items").SiebList("List").ActivateRow 0

	If sStockCheck <> "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Fulfillment").SiebPicklist("Stock Check").SiebSelect sStockCheck
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Fulfillment").SiebButton("Check Stock").SiebClick sPopUp
		CaptureSnapshot()
		Browser("index:=0").Page("index:=0").Sync
	End If
	
	If sStockCheckTextBox = "Warehouse" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Fulfillment").SiebPicklist("Stock Check").SiebSelect sStockCheckTextBox
		CaptureSnapshot()
		Browser("index:=0").Page("index:=0").Sync
	End If


	If Ucase(sReserve) = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Order Entry -Shipment").SiebApplet("Orders").SiebButton("Reserve").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
	End If
	
	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : OrdersDeliveryTab_fn
' 	Description : 
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function OrdersDeliveryTab_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [Delivery$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

			sValidation = adoData( "Validation")  & ""
			sDeliveryMethod = adoData( "DeliveryMethod")  & ""
			sPostcode = adoData( "Postcode")  & ""
			sPhoneNo = adoData( "PhoneNo")  & ""
			sValidation = adoData( "Validation")  & ""
			sReserve  = adoData( "Reserve")  & ""
            sPopUp = adoData("Popup") & ""
            sNoAdressValidation = adoData("NoAdressValidation") & ""
			 sSequence = adoData("Sequence") & ""
			 sNegativeVal = adoData("NegativeVal") & ""
			 sRemoveDelivery = adoData("RemoveDelivery") & ""
			 sDone = adoData("Done") & ""

	'
	'Flow

		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Ship To View (Sales)","L3"
		Browser("index:=0").Page("index:=0").Sync

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Delivery").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************
'To Click on Remove Delivery button.
			If sRemoveDelivery = "Y" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebButton("Remove Delivery").SiebClick False
			Browser("index:=0").Page("index:=0").Sync
			CaptureSnapshot()

		End If

'To Click on done button.
		If sDone = "Y" Then
			  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebButton("Done").SiebClick False
			  Browser("index:=0").Page("index:=0").Sync
			  CaptureSnapshot()
			  Exit Function
		End If

'To Validate Pack Instruction & Delivery Instruction field is not present under Delivery Tab.
			If sNegativeVal = "Y" Then
		
					sVal = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebPicklist("classname:=SiebPicklist","index:=0").GetRoProperty("uiname")
		
						If sVal<>"Pack instruction" Then
								AddVerificationInfoToResult "Info" , "Pack Instruction is not present in the Delivery details tab."
							Else 
								AddVerificationInfoToResult "Error" , "Pack Instruction is not present in the Delivery details tab."
								iModuleFail = 1
						 End If
		
						If sVal<>"Delivery Instruction" Then
					     		AddVerificationInfoToResult "Info" , "Delivery Instruction is not present in the Delivery details tab."
						  Else 
							   AddVerificationInfoToResult "Error" , "Delivery Instruction is not present in the Delivery details tab."
						    	iModuleFail = 1
						End If
					
					CaptureScrenshot()
				Exit Function
			End If

'To validate Text Box
			If sSequence = "Y" Then
		
					strArray = "Postcode;House no.;Flat;County;Delivery Email;Phone no.;First name;Store address;Postcode;County;Address line 1;Town/City;Pack instructions;Last name;Store code;Town/City;House/Flat name;Store name"
					
					strArraySplit = split(strArray,";")
			
					sSbTxtCnt = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").GetClassCount("SiebText")
													
					For loopVar = 0 to sSbTxtCnt-1
							sVal = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebText("classname:=SiebText","index:=" & loopVar).GetRoProperty("uiname")
		
							If Instr(sVal,strArraySplit(loopVar)) > 0 Then
									AddVerificationInfoToResult "Info" , sVal&"  :is sequentially present as Expected "
								Else 
									AddVerificationInfoToResult "Error" ,sVal& ":is not sequentially present as Expected"
									iModuleFail = 1
							End If
					Next
					CaptureScrenshot()
		
		
		'To validate for checkbox, calendar, calculator and picklist
						sExists = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebCheckbox("Deliver to store").Exist(5)
						If sExists = "True" Then
							AddVerificationInfoToResult "Info" , "Deliver to store is present in the tab"
							Else
							AddVerificationInfoToResult "Error" , "Deliver to store is not  present in the tab"
							iModuleFail = 1
							CaptureScrenshot()
						End If
		
						sExists = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebCalculator("Delivery cost").Exist(5)
							If sExists = "True" Then
							AddVerificationInfoToResult "Info" , "Delivery cost is present in the tab"
							Else
							AddVerificationInfoToResult "Error" , "Delivery cost is not  present in the tab"
							iModuleFail = 1
							CaptureScrenshot()
						End If
		
					sExists = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebPicklist("Delivery method").Exist(5)
							If sExists = "True" Then
								AddVerificationInfoToResult "Info" , "Delivery method is present in the tab"
							Else
								AddVerificationInfoToResult "Error" , "Delivery method is not  present in the tab"
								iModuleFail = 1
								CaptureScrenshot()
							End If
		
				sExists = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebCalendar("Delivery date").Exist(5)
					If sExists = "True" Then
					AddVerificationInfoToResult "Info" , "Delivery date is present in the tab"
					Else
					AddVerificationInfoToResult "Error" , "Delivery date is not  present in the tab"
					iModuleFail = 1
					CaptureScrenshot()
				End If
			Exit Function
		
		End If


        If (instr(sDeliveryMethod,"DPD STB Delivery")>0) Then
					If   sNoAdressValidation= "Yes"Then
				
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebPicklist("Delivery method").SiebSelect sDeliveryMethod
					CaptureSnapshot()
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebButton("Done").SiebClick sPopUp
				Exit Function
				End If
		End If

		If sNoAdressValidation= "Yes"  Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebButton("Done").SiebClick sPopUp
			CaptureSnapshot()
			Exit Function
		End If


		'Add Addresslin1
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebText("Address line 1").OpenPopup
		Browser("index:=0").Page("index:=0").Sync

		If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("VF Delivery Address Pick").Exist (5) Then
				CaptureSnapshot()
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("VF Delivery Address Pick").SiebButton("OK").Click
		Else
					AddVerificationInfoToResult "Fail","Address Pick Applet not enabled"
					iModuleFail = 1		
					Exit Function
		End If

		
		CaptureSnapshot()
	'Add Delivery Method
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebPicklist("Delivery method").SiebSelect sDeliveryMethod
	'Add Phone Num
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebText("Phone no.").SiebSetText sPhoneNo


			If (weekday(Date)) <> 6 Then
				AddVerificationInfoToResult "Pass","Oops ,Its not Fridayyy!!!!"
				   DateRequested = Date		
					DateRequested = DateAdd("d", 1, DateRequested)
'					daydt =Day(DateRequested)
					MN = Left( MonthName(Month(DateRequested)),3)
'					yr = Year(DateRequested)	
					DateRequested = day(DateRequested)   & "/" & MN & "/" &  year(DateRequested)				
					sDeliveryDate = DateRequested
				
			Else	
'					AddVerificationInfoToResult "Pass","Yayy ,Its Fridayyy!!!!"
'					'MN =Left( MonthName(Month(Date)),3)
'					DateRequested = Date					
'					DateRequested = day(DateRequested) + 3 & "/" & MN & "/" &  year(DateRequested)				
'					sDeliveryDate = DateRequested

					AddVerificationInfoToResult "Pass","Yayy ,Its Fridayyy!!!!"
					DateRequested = Date		
					DateRequested = DateAdd("d", 3, DateRequested)
'					daydt =Day(DateRequested)
					MN = Left( MonthName(Month(DateRequested)),3)
'					yr = Year(DateRequested)	
					DateRequested = day(DateRequested) & "/" & MN & "/" &  year(DateRequested)				
					sDeliveryDate = DateRequested
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebCalendar("Delivery date").SetText sDeliveryDate
			
			End If

			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebButton("Done").SiebClick False


	If UCase(sValidation) = "YES" Then
			
					sExchange = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebPicklist("Delivery method").IsExists("Exchange*")
					sReturns = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebPicklist("Delivery method").IsExists("Return*")
					CaptureSnapshot()
					sPostCode = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebText("Postcode").GetROProperty("isenabled")
				
						If sExchange and sExchange Then
				
							AddVerificationInfoToResult "Fail","Exchange and Returns delivery methods are present"
							iModuleFail = 1	
						Else
							AddVerificationInfoToResult "Pass","Exchange and Returns delivery methods are NOT present"			
						End If
				
						If sPostCode = "False"  Then
				
							AddVerificationInfoToResult "Pass","Postcode is NOT Editable"
						Else
							AddVerificationInfoToResult "Fail","Postcode is  Editable"
							iModuleFail = 1			
						End If
		End If

'		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************

		If  UCase(sReserve) = "YES" Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Reserve").SiebClick False
		End If
		
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : FLModifyOrderFormEntry_fn
' 	Description : This function log into Siebel Application
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function FLModifyOrderFormEntry_fn()

	'sOR
Dim adoData	 

    strSQL = "SELECT * FROM [FLModifyOrderFormEntry$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	sDisconnectionReason = adoData( "DisconnectionReason")  & ""
	sReturnReason =  adoData( "ReturnReason")  & ""
	sJourney = adoData( "Journey")  & ""
	sPopUp = adoData( "Popup")  & ""	
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If


	'Flow
	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")

	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If

	Browser("index:=0").Page("index:=0").Sync

'    '*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebCheckbox("Terms and conditions").SiebSetCheckbox "ON"

    If sDisconnectionReason <> "" Then
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Disconnection reason").Exist(2)=False Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("ToggleLayout").Click
		End If
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Disconnection reason").SiebSelect sDisconnectionReason
	End If

	
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Verify").SiebClick sPopUp ' clicking on Verfiy Button

	Browser("index:=0").Page("index:=0").Sync

	

	Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items")


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


	sPrimaryFirstName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("First Name").SiebGetRoProperty("text")
	If sPrimaryFirstName = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("First Name").SiebSetText "FIXED"
	End If

	sAlternateFirstName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("First Name_2").SiebGetRoProperty("text")
	If sAlternateFirstName = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("First Name_2").SiebSetText "FIXED1"
	End If

	sPrimaryLastName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Last Name").SiebGetRoProperty("text")
	If sPrimaryLastName = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Last Name").SiebSetText "LINE"
	End If

	sAlternateLastName = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Last Name_2").SiebGetRoProperty("text")
	If sAlternateLastName = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Last Name_2").SiebSetText "LINE1"
	End If

	sPrimaryPhoneNo = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Phone no.").SiebGetRoProperty("text")
	If sPrimaryPhoneNo = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Phone no.").SiebSetText "9999995642"
	End If

	sAlternatePhoneNo = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Phone no._2").SiebGetRoProperty("text")
	If sAlternatePhoneNo = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Phone no._2").SiebSetText "9999995641"
	End If

	sPrimaryAltPhoneNo = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Alt Phone no.").SiebGetRoProperty("text")
	If sPrimaryAltPhoneNo = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Alt Phone no.").SiebSetText "9999995643"
	End If

	sAlternateAltPhoneNo = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Alt Phone no._2").SiebGetRoProperty("text")
	If sAlternateAltPhoneNo = "" Then
		 SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Alt Phone no._2").SiebSetText "9999995644"
	End If

	sPrimaryEmail  = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Email").SiebGetRoProperty("text")
	If sPrimaryEmail = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Email").SiebSetText "abc@sqcmail.uk"
	End If
	
	sAlternateEmail = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Email_2").SiebGetRoProperty("text")
	If sAlternateEmail = "" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebText("Email_2").SiebSetText "abc1@sqcmail.uk"
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Verify").SiebClick sPopUp ' clicking on Verfiy Button

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Submit").SiebClick sPopUp

	CaptureSnapshot()
   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : Econfig_TV_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function Econfig_TV_fn()

'Wait 5
Dim adoData	 

    strSQL = "SELECT * FROM [Econfig$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Econfig.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow
		Do while Not adoData.Eof
			sLink = adoData( "Link")  & ""
			sItem = adoData( "Item")  & ""

			sCheck = adoData( "Check")  & ""

			sList = adoData( "List")  & ""
			sCustomise= adoData( "Customise")  & ""
			'sCustomiseItem= adoData( "CustomiseItem")  & ""
			sTextValue= adoData( "TextValue")  & ""
			sAddItem= adoData( "AddItem")  & ""
			sAction = adoData( "Action")  & ""
			sCatalog = adoData( "Catalog")  & ""

			Browser("index:=0").Page("index:=0").Sync

		    If  Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Exist(3) Then
				 Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebElement("OK").Click
			End If

			
			If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Exist(5) Then
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("Proceed").Click
			End If

			Browser("index:=0").Page("index:=0").Sync

			If sLink<>"" Then				
			Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").Link("text:="&sLink,"html tag:=A").WebClick
'			Wait 5
			End If

			CaptureSnapshot()

			If sCheck<>"" Then

'				If sItem="A bar" or sItem="O bar" or sItem = "Credit Alert A Bar" or sItem="Fraud" or sItem="Stolen Bar" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//a[text()='"&sItem&"']/../..").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'				Else
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebCheckBox("html tag:=INPUT","index:=0").WebSet sCheck
'				End If
				AddVerificationInfoToResult  "Info" ,"Action :" & sCheck & " is done on product " &sItem
			End If

			If sCatalog <> "" Then
				If Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("There is an error with").Exist(3) Then
					AddVerificationInfoToResult "Info" , "Pop up message is displayed successfully after selecting " & sItem & " for Prepaid Account."
					CaptureSnapshot()
					Exit Function
				Else
					AddVerificationInfoToResult "Error" , "Pop up message is not displayed successfully."
					iModuleFail = 1
					Exit Function
				End If
			End If

			If sCustomise="Y" Then
'				If sLink="Bars" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebTable("column names:=.*Customise").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
'					Wait 5
'				else
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Image("file name:=icon_configure\.gif").WebClick
'				End If
				
				'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sCustomiseItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sCustomiseValue
			End If			

			If sList<>"" Then
            '				If sItem="Item" Then
					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//select/option[text()='"&sList&"']/..").WebSelect sList
'				ElseIf sItem ="AreaCode" Then
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebList("all items:=.*"&sList&".*").WebSelect sList
'				Else
'					Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebList("xpath:=//*[text()='"&sItem&"']/../..//select/option[text()='"&sList&"']/..").WebSelect sList			
'				End If
			End If

			If sAddItem<>"" Then
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").WebEdit("html tag:=INPUT").Set sAddItem

				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("xpath:=//select/option[text()='"&sList&"']/../../..").Link("text:=Add Item").WebClick
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT","index:=0").Set sAddItem
'				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").Link("text:=Add Item").WebClick
			End If	


If  sTextValue<>""Then
	If sItem ="MSISDN" Then
				sTextValue = DictionaryTest_G.Item("ACCNTMSISDN")
'                sTextValue = "447387945790"
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue
	Else
				Browser("Siebel Call Center").Page("Siebel Call Center").Frame("CfgMainFrame Frame").WebElement("html tag:=TR","innertext:="&sItem&".*","index:=0").WebEdit("html tag:=INPUT").WebSet sTextValue

			End If

End If
			
			adoData.movenext
		
	Loop

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : OrderLineDetailVerify_fn
' 	Description : This function verify the required fields in Order Line detail section
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function OrderLineDetailVerify_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [OrderLineDetail$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

    sExpectedComment = 	adoData( "ExpectedComment")  & ""

	'Flow

		Browser("index:=0").Page("index:=0").Sync

		sActualComment = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebTextArea("Eligibility comments").GetROProperty("text")

		If Ucase(sActualComment) = Ucase(sExpectedComment)  Then

				AddVerificationInfoToResult "Pass","Comment is as expected"
		 Else
				AddVerificationInfoToResult "Fail","Comment is not as expected"
						iModuleFail = 1
		End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function



'#################################################################################################
' 	Function Name : SelectBillingProfle_fn
' 	Description : This function Selects reuired billing profile
'   Created By :  Pushkar Vasekar
'	Creation Date :        
'##################################################################################################
Public Function SelectBillingProfle_fn()

Dim adoData	  
    strSQL = "SELECT * FROM [SelectBillingProfle$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebText("Billing profile").OpenPopup
Browser("index:=0").Page("index:=0").Sync

Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile")
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
					adoData.MoveNext
			Loop ''

		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Billing profile").SiebButton("OK").SiebClick False

			CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


	End Function
 

'#################################################################################################
' 	Function Name : Delivery_Postcode_fn
' 	Description : This function is used to click on delivery tab and to chech if postcode is populated on its own after slecting address
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function Delivery_Postcode_fn()

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Ship To View (Sales)","L3" ' clicking on delivery tab on orders page


''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Delivery").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************
	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebText("Address line 1").OpenPopup
		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("VF Delivery Address Pick").Exist(5) Then
			CaptureSnapshot()
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("VF Delivery Address Pick").SiebButton("OK").SiebClick False
		Else
			AddVerificationInfoToResult  "Error" , "Open Pop up is not clicked for Addressline 1."
			iModuleFail = 1
			Exit Function
		End If
	Else
		AddVerificationInfoToResult  "Error" , "Delivery tab is not clicked successfully."
		iModuleFail = 1
		Exit Function
	End If

	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebText("Postcode").Exist(5) Then
		sPostCode = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Shipment").SiebApplet("Shipment Details").SiebText("Postcode").GetRoProperty("text")
	Else
		AddVerificationInfoToResult  "Error" , "Delivery Page is not displayed successfully."
		iModuleFail = 1
		Exit Function
	End If

	If sPostCode <> "" Then
		AddVerificationInfoToResult  "Info" , "Post code is populated automatically after selecting address from address line 1 and is : "& sPostCode
		CaptureSnapshot()
	Else
		AddVerificationInfoToResult  "Error" , "Post code is not populated automatically after selecting address from address line."
		iModuleFail = 1
	End If


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function
 

	'#################################################################################################
' 	Function Name : RetrieveBulkMSISDN_fn
' 	Description : This function makes number of copies of the selected promotion and retrieves MSISDNs for all
'   Created By :  Pushkar Vasekar	
'	Creation Date :        
'##################################################################################################
Public Function RetrieveBulkMSISDN_fn()

'Dim adoData	  
'    strSQL = "SELECT * FROM [Order_Payment$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If

''*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************
	CaptureSnapshot()
	Browser("index:=0").Page("index:=0").Sync	

SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebButton("Copy").Click
Browser("index:=0").Page("index:=0").Sync

SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Bulk Copy").SiebCalculator("Quantity").SetText "3"

SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Bulk Copy").SiebButton("Go").SiebClick False
Browser("index:=0").Page("index:=0").Sync

SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebButton("Retrieve MSISDNs").SiebClick False
Browser("index:=0").Page("index:=0").Sync

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
' 	Function Name :OrderStatusCheck_fn
' 	Description : This function log into Siebel Application
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function OrderStatusCheck_fn()


	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	Browser("index:=0").Page("index:=0").Sync
	activeView=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").getROProperty("activeview")
	If activeView<>"Order Entry - Line Items Detail View (Sales)" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	End If

'	'*****************Web Element********************
'Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").Link("Line Items").Click
'Browser("index:=0").Page("index:=0").Sync
''*****************Web Element********************

	Browser("index:=0").Page("index:=0").Sync
	
				sStatus=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Status").SiebGetRoProperty("activeitem")
				i = 0
		
				Do while sStatus <> "Complete" and i <=50			
					wait (5)			
					SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebMenu("Menu").Select "ExecuteQuery"			
					sStatus=SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebPicklist("Status").SiebGetRoProperty("activeitem")
			
					If sStatus = "Complete" Then
						Exit Do
					End If			
					i = i+1			
				Loop
		
				If   sStatus <> "Complete"Then
						  iModuleFail = 1
							AddVerificationInfoToResult  "Error" , "Order Sattaus is not Complete"
				End If
	
	CaptureSnapshot()
   
			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If


		End Function
'#################################################################################################
' 	Function Name : OrderNumberSelection_fn
' 	Description :  This function is used to select the order number from orders tab, check the order status and click on customer account button.
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function OrderListVerification_fn()


	Dim adoData	  

    strSQL = "SELECT * FROM [OrderListVerification$] WHERE  [RowNo]=" & iRowNo
      Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

                'sOR
                call SetObjRepository ("Orders",sProductDir & "Resources\")

               ' sGotoOrder = adoData( "GotoOrder")  & ""             
                sLastNameValidation =  adoData("LastNameValidation")  & ""

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sOrderNumber =  DictionaryTest_G.Item("OrderNo")
'sOrderNumber=	"SBL-1000000000390363"

'	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Sales Order Screen" ' clicking on Orders tab

	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Agreement Screen"
	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Sales Order Screen"
	Browser("index:=0").Page("index:=0").Sync
	


	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - All Orders View (Sales)","L2"

			If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").Exist(5) Then
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebButton("Query").SiebClick False
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebList("List").SiebText("Order no.").SiebSetText sOrderNumber
				SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebButton("Go").SiebClick False
				Browser("index:=0").Page("index:=0").Sync
	
				If 	(sLastNameValidation="NotReadOnly") Then
						SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebList("List").SiebText("Last name").OpenPopUp 
						If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Pick Contact").Exist(2)Then
							rowCnt =SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Pick Contact").SiebList("List").SiebListRowCount
								CaptureSnapshot
				
													If rowCnt <> 0 Then
														AddVerificationInfoToResult "Pass","Rows are displayed. Last Name is editable.."
													Else
														iModuleFail = 1
														AddVerificationInfoToResult "Fail","No rows displayed. Last Name is not editable."
														Exit Function
													End If
												SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Pick Contact").SiebButton("OK").Click								
						End If 
				End If

				If 	(sLastNameValidation="ReadOnly") Then
					strEnable = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("All Orders").SiebApplet("Orders").SiebList("List").SiebText("Last name").getROProperty("isenabled")
					CaptureSnapshot
					If  strEnable = False Then
								AddVerificationInfoToResult "Pass","Last Name is not editable.."
						Else
								iModuleFail = 1
								AddVerificationInfoToResult "Fail","Last Name is editable."
						Exit Function
					End If
				End If


			Else
				AddVerificationInfoToResult  "Fail" ,"Orders Page is not displayed successfully."
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
' 	Function Name : DeleteProduct_fn
' 	Description :  This function is used to delete the products.
'   Created By :  Vishwa
'	Creation Date :        
'##################################################################################################
Public Function DeleteProduct_fn()


	'Flow


	Dim adoData	  
    strSQL = "SELECT * FROM [AddProduct$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sDeleteProduct = adoData( "DeleteProduct")  & ""
	

	Browser("index:=0").Page("index:=0").Sync

	'Set obj=Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebTable("column names:="&sDeleteProduct&".*","html tag:=TABLE")
	Set obj=Browser("Siebel Call Center").Page("Siebel Call Center").Frame("View Frame").WebTable("column names:=Postpay Handsets: Postpay Handsets Add","html tag:=TABLE")

	If obj.Exist(7) Then
		'obj.Link("innertext:=Add").WebClick
		obj.Image("alt:=Delete","repositoryname:=Delete.32").WebClick
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("SiebView").SiebApplet("Edit Package").SiebButton("Done").SiebClick False
	Else
		AddVerificationInfoToResult  "Fail" ,"Product: " & sDeleteProduct & " is not available on page."
		iModuleFail = 1

		

	End If

	End Function
 

'#################################################################################################
' 	Function Name : ClickReviseButton_fn
' 	Description : This function is used to click on Revise Button.
'   Created By :  Vishwa
'	Creation Date :        
'##################################################################################################
Public Function ClickReviseButton_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

		If  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Revise").Exist(5) Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Revise").SiebClick False
			Browser("index:=0").Page("index:=0").Sync
			CaptureSnapshot()
	End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : VerifyOrdersAboutView_fn
' 	Description : This function is used to click on Help Button and Verify the About View.
'   Created By :  Zeba
'	Creation Date :        
'##################################################################################################
Public Function VerifyOrdersAboutView_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [BillingProfile$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Account.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Sales Order Screen"
		Browser("index:=0").Page("index:=0").Sync
		CaptureSnapshot()

		SiebApplication("Siebel Call Center").SiebMenu("Menu").Select "Help\\Help - About View"
		Browser("index:=0").Page("index:=0").Sync
		CaptureSnapshot()




		If Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").Exist(5) then
					AddVerificationInfoToResult "PASS","About View Applet Exists as Expected."
					CaptureSnapshot()
					Browser("SiebWebPopupWindow").Page("SiebWebPopupWindow").Frame("_swepopcontent Frame").SblButton("OK").Click
        Else
				   AddVerificationInfoToResult "Fail","About View Applet does not Exists"
					iModuleFail = 1
					Exit Function
		End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : VerifyColumnsOrderView_fn
' 	Description : This function is used to Enter Installed id and Verify Specific Columns
'   Created By :  Zeba
'	Creation Date :        
'##################################################################################################
Public Function VerifyColumnsOrderView_fn()

	'Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [VerifyColumnsOrderView$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Order.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoScreen "Sales Order Screen"
		Browser("index:=0").Page("index:=0").Sync
		CaptureSnapshot()

		SiebApplication("Siebel Call Center").SiebPageTabs("PageTabs").GotoView "VF MSISDN SearchOrder View"
		Browser("index:=0").Page("index:=0").Sync
		CaptureSnapshot()

		sInstalledID =DictionaryTest_G("ACCNTMSISDN") 
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("SearchOrder").SiebApplet("Installed ID").SiebText("Installed ID").SetText sInstalledID
		CaptureSnapshot()
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("SearchOrder").SiebApplet("Installed ID").SiebButton("Go").Click
		CaptureSnapshot()

		Set objApplet =  SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("SearchOrder").SiebApplet("Orders Line Items")
		colCnt = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("SearchOrder").SiebApplet("Orders Line Items").SiebList("List").ColumnsCount

					For loopVar = 0 to colCnt
						sReposName = objApplet.SiebList("List").GetColumnRepositoryNameByIndex(loopVar)
						sUIName = objApplet.SiebList("List").GetColumnUIName(sReposName)
						If sUIName = "Order Number " Then
							val1 = loopVar
						End If
				
						If sUIName = "Order Type" Then
							val2 = loopVar
						End If
				
						If sUIName = "Order Subtype" Then
							val3 = loopVar
						End If
					
					Next
'				
					If Cint(val3) = Cint(val2) + 1 Then
						AddVerificationInfoToResult  "Info" ,"Order Subtype is beside Order Type "
						CaptureSnapshot()
					Else
						AddVerificationInfoToResult  "Error" ,"Order Subtype is not beside Order Type "
						iModuleFail = 1
					End If
				
					If Cint(val2) = Cint(val1) + 1 Then
						AddVerificationInfoToResult  "Info" ,"Order Type is beside Order Subtype "
						CaptureSnapshot()
					Else
						AddVerificationInfoToResult  "Error" ,"Order Type is not beside Order Subtype "
						iModuleFail = 1
					End If
				
					If Cint(val1) = Cint(val2) - 1 Then
						AddVerificationInfoToResult  "Info" ,"Order Number is beside Order Subtype "
						CaptureSnapshot()
					Else
						AddVerificationInfoToResult  "Error" ,"Order Number is not beside Order Subtype "
						iModuleFail = 1
					End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function
