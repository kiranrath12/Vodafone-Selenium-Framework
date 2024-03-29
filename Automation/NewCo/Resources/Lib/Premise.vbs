	'#################################################################################################
' 	Function Name : CustomerAccount_fn
' 	Description : This function is used to used to click on line items tab and Customer Account button.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CustomerAccount_fn()

    	
	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

    SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3" ' clicking on line items tab
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Orders").SiebButton("Customer Account").SiebClick False ' clicking on customer account button on Orders page
  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : CreateNewPremise_fn
' 	Description : This function is used to used to create new premise in Premise tab.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function CreateNewPremise_fn()

	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Address$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\DataProxy.xls",strSQL)
	
	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sCustomerType = adoData("CustomerType") & ""

	sAddress = DictionaryTest_G.Item("Address")

	sPostCode = Trim(Split(sAddress,"|")(0))
	sHouseNumber = Split(sAddress,"|")(1)

	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF CUT Account Premise Service Point View","L3" ' clicking on Premise tab

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("All Premises").SiebButton("New Premise").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebList("List").SiebText("Post Code").SiebSetText sPostCode
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebList("List").SiebText("House number").SiebSetText sHouseNumber
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebButton("Go").SiebClick False
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebButton("Create Premise").SiebClick False

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Query").SiebClick False ' clicking on Query button
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Go").SiebClick False

	Browser("index:=0").Page("index:=0").Sync
	rwCnt = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("All Premises").SiebList("List").SiebListRowCount

	If rwCnt <> 0 Then
		AddVerificationInfoToResult  "Info" , "Premise is created successfully."
	Else
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("All Premises").SiebButton("New Premise").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebList("List").SiebText("Post Code").SiebSetText sPostCode
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebList("List").SiebText("House number").SiebSetText sHouseNumber
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebButton("Go").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebButton("Create Premise").SiebClick False

		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Query").SiebClick False ' clicking on Query button
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Go").SiebClick False

	End If

	If rwCnt <> 0 Then
		AddVerificationInfoToResult  "Info" , "Premise is created successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Premise is not created successfully."
		iModuleFail = 1
		Exit Function
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebPicklist("Customer Type").SiebSelect sCustomerType

  
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : BookAppointment_fn
' 	Description :  This function is used to book appointment and Reserve next number on line items page by selecting Fixed Line Service
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function BookAppointment_fn()


Dim adoData	  
    strSQL = "SELECT * FROM [FilterPromotion$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Premise.xls",strSQL)

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	sSelectItem = adoData( "SelectItem")  & ""
	sRowNum = adoData( "RowNum")  & ""
	sBookAppointment = adoData( "BookAppointment")  & ""
	sReserveNextNumber = adoData( "ReserveNextNumber")  & ""

	sPopUp3 = adoData( "Popup3")  & ""	'2nd Popup
	If Ucase(sPopUp3)="FALSE" Then
		sPopUp3 =Cbool(sPopUp3)
	End If

	sRowNum1 = Cint(sRowNum)

	Browser("index:=0").Page("index:=0").Sync
	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebScreenViews("ScreenViews").Goto "Order Entry - Line Items Detail View (Sales)","L3"
	Browser("index:=0").Page("index:=0").Sync

	If sBookAppointment = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectProductRow sSelectItem,0
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebButton("BookAppointment").SiebClick sPopUp3 ' Book Appointment after selecting fixed line service

		Browser("index:=0").Page("index:=0").Sync
	
		On Error Resume Next
		Recovery.Enabled=False
	
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Book Appointment").SiebList("List").ActivateRow sRowNum1
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Book Appointment").SiebButton("Book Appointment").SiebClick sPopUp3

		If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Book Appointment").SiebButton("Book Appointment").Exist(2) Then
			SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Book Appointment").SiebButton("Book Appointment").SiebClick False
		End If
	
		Err.Clear
		Recovery.Enabled=True
	End If

	Browser("index:=0").Page("index:=0").Sync

	If sReserveNextNumber = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectProductRow sSelectItem,0
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebButton("Reserve Next Number").SiebClick sPopUp3 ' Reserve Next Number after selecting Fixed line service
	End If

   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function


'#################################################################################################
' 	Function Name : OrderLineDetail_fn
' 	Description :  This function is used to enter value of entry type in line items page after selecting fixed line service.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function OrderLineDetail_fn()


Dim adoData	  
	 strSQL = "SELECT * FROM [FilterPromotion$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Premise.xls",strSQL)

	sSelectItem = adoData( "SelectItem")  & ""

	'sOR
	call SetObjRepository ("Orders",sProductDir & "Resources\")

	'Flow

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SelectProductRow sSelectItem,0   ' selecting Vodafone connect router

	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Detail").SiebPicklist("EntryType").SiebSelect "Ordinary" ' Inserting Entry Type

   
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function



'#################################################################################################
' 	Function Name : FilterPromotion_fn
' 	Description :  This function is used to select the correct promotion after clicking on Service Availability check.
'   Created By :  Tarun/Ankit
'	Creation Date :        
'##################################################################################################
	Public Function FilterPromotion_fn()
	
	
	Dim adoData	  
		strSQL = "SELECT * FROM [FilterPromotion$] WHERE  [RowNo]=" & iRowNo
		Set adoData = ExecExcQry(sProductDir & "Data\Premise.xls",strSQL)

		sDirectoryNumber = DictionaryTest_G.Item("RetnNo")
		sJourneyType = adoData( "JourneyType")  & ""
		sLineInfo = adoData( "LineInfo")  & ""
		sAction = adoData( "Action")  & ""
		sNumRetn = adoData( "NumRetn")  & ""
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""

		sPopUp = adoData( "Popup")  & ""	' Ist Popup
		If Ucase(sPopUp)="NO" Then
			sPopUp="FALSE"
			sPopUp=Cbool(sPopUp)
		End If

		
		sPopUp2 = adoData( "Popup2")  & ""	'2nd Popup
		If Ucase(sPopUp2)="NO" Then
			sPopUp2="FALSE"
			sPopUp2=Cbool(sPopUp2)
		End If

'		Set oDict = CreateObject("Scripting.Dictionary")
'
'		If instr(sLocateColValue,"Promotion")>0 Then
'			sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
'		End If
'		arr1 = Split (sLocateCol,"|")
'		arr2 = Split (sLocateColValue,"|")
'		For i=0 to Ubound (arr1)
'			oDict.Add arr1(i),arr2(i)
'
'		Next

		'sOR
		call SetObjRepository ("Account",sProductDir & "Resources\")


		'Flow 
'
	If sNumRetn = "Y" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebText("Directory Number").SiebSetText sDirectoryNumber ' setting Directory number
	End If

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Service Availability Check").SiebClick sPopUp ' clicking on service availability check button

	Dim vMatchRowNium 
	Dim vNumOfResults
	 Set obj=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebList("List")
	 vNumOfResults = obj.SiebListRowCount

	 If  vNumOfResults = 0 Then 
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Query").SiebClick False ' clicking on Query button t o refresh tha page
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Go").SiebClick False
	End If

	 vNumOfResults = obj.SiebListRowCount

	 If  vNumOfResults = 0 Then ' again clicking on Service availabilty check if promotion list does not appear in 1st click
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Service Availability Check").SiebClick sPopUp ' clicking on service availability check button
		 Set obj=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebList("List")
		  vNumOfResults = obj.SiebListRowCount
	End If

	
	If  vNumOfResults = 0 Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "No rows displayed after clicking on Service Availability check."
		Exit Function
	End If
	
	Set obj1 = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results")
	Set obj=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebList("List")	  
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebButton("ToggleListRowCount").SiebClick False ' Expanding Availability check results list
	vNumOfResults = obj.SiebListRowCount
	 Dim  j
	Dim  Match,flag
	Match="N"
	flag="N"
	j=0
	index = 0

	For i=1 to 2	

		Dim LineInfoAct

		j=j+1

		
		res=LocateColumns (obj1,sLocateCol,sLocateColValue,index)
		index = index + 1	
		If Cstr(res)<>"True" Then
			 iModuleFail = 1
			AddVerificationInfoToResult  "Error" , "Promotion not found in the list."
			Exit Function
		End If

		checkLineInfo="N"
		If sAction ="Use Existing Line" Then
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebButton("Use Existing Line").getROProperty("isenabled")=True Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebButton("Use Existing Line").SiebClick sPopUp2
				checkLineInfo="Y"
			End If

		Elseif sAction ="Install New Line" Then
				Set menu = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebMenu("Menu")
				menuItem=menu.GetRepositoryName("Install New Line")
				On Error Resume Next
				Recovery.Enabled=False
				If menu.IsEnabled(menuItem) Then
					menu.Select(menuItem)
					checkLineInfo="Y"		

				If Browser("Siebel Call Center").Dialog("Message from webpage").Static("New Line should only be").Exist(5) Then ' Clicking on Pop up which appears after selecting Install new line
					Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
				End If
				
				Err.clear

				Recovery.Enabled=True

'				Set vDictionary = CreateObject("Scripting.Dictionary")
'				CheckForPopup vDictionary,sPopUp3			
				End If
			
		End If

		If checkLineInfo="Y" Then
			
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Query").SiebClick False ' clicking on Query button
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Go").SiebClick False

			LineInfoAct = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Service Point Detail").SiebText("Line Info").SiebGetRoProperty("text") ' Capturing Line Info value from service point detail


			If Ucase(LineInfoAct)=Ucase(sLineInfo) Then
				AddVerificationInfoToResult  "Info" , "Line Info: " & LineInfoAct & "  is same as  the expected value: " & sLineInfo
				Match="Y"
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Service Points").SiebButton("Create Order").SiebClick False
				Exit For

			End If 
		End If

		If Match="N" and checkLineInfo="Y" Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , "Line Info: " & LineInfoAct & "  is not matching to the expected value: " & sLineInfo
			Exit Function
		End If
	Next
	
	If Match="N" Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Line Info: " & LineInfoAct & "  is not matching to the expected value: " & sLineInfo
	End If
	
	
	   
		If Err.Number <> 0 Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , Err.Description
		End If


End Function




'#################################################################################################
' 	Function Name : PremiseMove_fn
' 	Description :  This function is used to click on Profile tab and create new billing profile.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function PremiseMove_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [PremiseMove$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Premise.xls",strSQL)

	sAddress = DictionaryTest_G.Item("Address")
	sPostCode = Trim(Split(sAddress,"|")(0))
	sHouseNumber = Split(sAddress,"|")(1)

	sClickNewPremise = adoData( "ClickNewPremise")  & ""
	sMoveType = adoData( "MoveType")  & ""
	sRetainNumber = adoData( "RetainNumber")  & ""
	sEnterDetails = adoData( "EnterDetails")  & ""
	sMSISDNLen = Len(DictionaryTest_G.Item("ACCNTMSISDN"))
	sMSISDN = "0" & Right(DictionaryTest_G.Item("ACCNTMSISDN"),sMSISDNLen-2)

	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	Browser("index:=0").Page("index:=0").Sync

	'Flow

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").GetRoProperty("ActiveScreen") <>  "VF CUT Account Premise Service Point View" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebScreenViews("ScreenViews").Goto "VF CUT Account Premise Service Point View","L3"
	End If

	Browser("index:=0").Page("index:=0").Sync

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("All Premises").SiebList("List").ActivateRow 0

    Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("All Premises")


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

			If sValue = "ConsumerCustomerType" Then
				sValue=	Replace(sValue,"ConsumerCustomerType","Individual")
			ElseIf sValue = "BusinessCustomerType" Then
				sValue=	Replace(sValue,"BusinessCustomerType","Business")
			End If

			UpdateSiebList objApplet,sUIName,sValue

			If instr (sUIName, "CaptureTextValue1") > 0 AND instr (sUIName, "Postcode") > 0 Then
				If DictionaryTest_G.Exists("Postcode") Then
					DictionaryTest_G.Item("Postcode")=sValue
				else
					DictionaryTest_G.add "Postcode",sValue
				End If
				AddVerificationInfoToResult  "Info" , "Postcode Value is : " & DictionaryTest_G.Item("Postcode")
			End If

			If instr (sUIName, "CaptureTextValue1") > 0 AND instr (sUIName, "Customer Type") > 0 Then
				If DictionaryTest_G.Exists("CustomerType") Then
					DictionaryTest_G.Item("CustomerType")=sValue
				else
					DictionaryTest_G.add "CustomerType",sValue
				End If
				AddVerificationInfoToResult  "Info" , "Customer Type is : " & DictionaryTest_G.Item("CustomerType")
			End If

		End If

		adoData.MoveNext
	Loop

	If sClickNewPremise = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("All Premises").SiebButton("New Premise").SiebClick False
		If sRetainNumber = "Yes" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebList("List").SiebText("Post Code").SiebSetText DictionaryTest_G.Item("Postcode")
		Else
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebList("List").SiebText("Post Code").SiebSetText sPostCode
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebList("List").SiebText("House number").SiebSetText sHouseNumber
		End If
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebButton("Go").SiebClick False
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("New Premise").SiebButton("Create Premise").SiebClick False
	
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Query").SiebClick False ' clicking on Query button
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Go").SiebClick False
		Browser("index:=0").Page("index:=0").Sync
		CaptureSnapshot()
	End If


	If sEnterDetails = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebPicklist("Customer Type").SiebSelect DictionaryTest_G.Item("CustomerType")
		If  sRetainNumber = "Yes" Then
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebText("Directory Number").SiebSetText sMSISDN
		Else
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebText("Directory Number").SiebSetText " "
		End If
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebPicklist("Move Type").SiebSelect sMoveType
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebText("Move Out Postcode").SiebSetText DictionaryTest_G.Item("Postcode")
	End If


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : PremiseMoveFilterPromotion_fn
' 	Description :  This function is used to select the correct promotion after clicking on Service Availability check.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
	Public Function PremiseMoveFilterPromotion_fn()
	
	
	Dim adoData	  
		strSQL = "SELECT * FROM [FilterPromotion$] WHERE  [RowNo]=" & iRowNo
		Set adoData = ExecExcQry(sProductDir & "Data\Premise.xls",strSQL)

		sLineInfo = adoData( "LineInfo")  & ""
		sAction = adoData( "Action")  & ""
		sLocateCol = adoData( "LocateCol")  & ""
		sLocateColValue = adoData( "LocateColValue")  & ""

		sPopUp = adoData( "Popup")  & ""	' Ist Popup
		If Ucase(sPopUp)="FALSE" Then
			sPopUp=Cbool(sPopUp)
		End If

		
		sPopUp2 = adoData( "Popup2")  & ""	'2nd Popup
		If Ucase(sPopUp2)="FALSE" Then
			sPopUp2 =Cbool(sPopUp2)
		End If


		'sOR
		call SetObjRepository ("Account",sProductDir & "Resources\")


		'Flow 

	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Service Availability Check").SiebClick sPopUp ' clicking on service availability check button

	Dim vMatchRowNium 
	Dim vNumOfResults
	 Set obj=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebList("List")
	 vNumOfResults = obj.SiebListRowCount

	 If  vNumOfResults = 0 Then 
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Query").SiebClick False ' clicking on Query button t o refresh tha page
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Go").SiebClick False
	End If

	 vNumOfResults = obj.SiebListRowCount

	 If  vNumOfResults = 0 Then ' again clicking on Service availabilty check if promotion list does not appear in 1st click
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Service Availability Check").SiebClick sPopUp ' clicking on service availability check button
		 Set obj=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebList("List")
		  vNumOfResults = obj.SiebListRowCount
	End If

	
	If  vNumOfResults = 0 Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "No rows displayed after clicking on Service Availability check."
		Exit Function
	End If
	
	Set obj1 = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results")
	Set obj=SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebList("List")	  
	SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebButton("ToggleListRowCount").SiebClick False ' Expanding Availability check results list
	vNumOfResults = obj.SiebListRowCount
	 Dim  j
	Dim  Match,flag
	Match="N"
	flag="N"
	j=0
	index = 0

	For i=1 to 2	

		Dim LineInfoAct

		j=j+1

		
		res=LocateColumns (obj1,sLocateCol,sLocateColValue,index)
		index = index + 1	
		If Cstr(res)<>"True" Then
			 iModuleFail = 1
			AddVerificationInfoToResult  "Error" , "Promotion not found in the list."
			Exit Function
		End If

		checkLineInfo="N"
		If sAction ="Use Existing Line" Then
			If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebButton("Use Existing Line").getROProperty("isenabled")=True Then
				SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebButton("Use Existing Line").SiebClick sPopUp2
				checkLineInfo="Y"
			End If

		Elseif sAction ="Install New Line" Then
				Set menu = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Availability Check Results").SiebMenu("Menu")
				menuItem=menu.GetRepositoryName("Install New Line")
				On Error Resume Next
				Recovery.Enabled=False
				If menu.IsEnabled(menuItem) Then
					menu.Select(menuItem)
					checkLineInfo="Y"		

				If Browser("Siebel Call Center").Dialog("Message from webpage").Static("New Line should only be").Exist(5) Then ' Clicking on Pop up which appears after selecting Install new line
					Browser("Siebel Call Center").Dialog("Message from webpage").WinButton("OK").Click
				End If
				
				Err.clear

				Recovery.Enabled=True
		
				End If
			
		End If

		If checkLineInfo="Y" Then
			
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Query").SiebClick False ' clicking on Query button
			SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Premise").SiebButton("Go").SiebClick False

			LineInfoAct = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Address Profile").SiebApplet("Service Point Detail").SiebText("Line Info").SiebGetRoProperty("text") ' Capturing Line Info value from service point detail


			If Ucase(LineInfoAct)=Ucase(sLineInfo) Then
				AddVerificationInfoToResult  "Info" , "Line Info: " & LineInfoAct & "  is same as  the expected value: " & sLineInfo
				Match="Y"
				Exit For
			End If 
		End If

		If Match="N" and checkLineInfo="Y" Then
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , "Line Info: " & LineInfoAct & "  is not matching to the expected value: " & sLineInfo
			Exit Function
		End If
	Next
	
	If Match="N" Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Line Info: " & LineInfoAct & "  is not matching to the expected value: " & sLineInfo
	End If
	
	
	   
		If Err.Number <> 0 Then
				iModuleFail = 1
				AddVerificationInfoToResult  "Error" , Err.Description
		End If


End Function

'#################################################################################################
' 	Function Name : AccountPremises_fn
' 	Description :  This function is used to click on Profile tab and create new billing profile.
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function AccountPremises_fn()


	Dim adoData	  
    strSQL = "SELECT * FROM [AccountPremises$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Premise.xls",strSQL)


	'sOR
	call SetObjRepository ("Account",sProductDir & "Resources\")

	sClickOkButton = adoData( "ClickOkButton")  & ""

	'Flow

	Browser("index:=0").Page("index:=0").Sync

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Premises").Exist(10) Then
		AddVerificationInfoToResult  "Info" , "Applet is displayed successfully after selecting Premises Move from Menu items"
		CaptureSnapshot()
	Else
		AddVerificationInfoToResult  "Error" , "Applet is not displayed after selecting Premises Move from Menu items"
		iModuleFail = 1
		Exit Function
	End If

	MN =Left( MonthName(Month(Date)),3)
	DateRequested = Date + "7"				
	sNextDate = day(DateRequested) & "/" & Left( MonthName(Month(DateRequested)),3) & "/" &  year(DateRequested) 


    Set objApplet = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Premises")

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
			If sValue = "Date" Then
				sValue=	Replace(sValue,"Date",sNextDate)
			End If
			UpdateSiebList objApplet,sUIName,sValue
		End If

		adoData.MoveNext
	Loop

	If sClickOkButton = "Yes" Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Summary").SiebApplet("Account Premises").SiebButton("OK").SiebClick False
	End If

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function