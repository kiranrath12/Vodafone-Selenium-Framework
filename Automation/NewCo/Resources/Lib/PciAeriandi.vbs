'#################################################################################################
' 	Function Name : AeriandiSecureCallControlEntry_fn
' 	Description : This function enters details of creditcard and submits the application
'   Created By :  Tarun Bansal
'	Creation Date :        
'##################################################################################################
Public Function AeriandiSecureCallControlEntry_fn()
   	'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Aeriandi$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PCI.xls",strSQL)

	sCardNumberE7 = adoData( "CardNumberE7")  & ""	
	sCardNumberE4 = adoData( "CardNumberE4")  & ""
	sCardNumberC2 = adoData( "CardNumberC2")  & ""
	sExpiryDate = adoData( "ExpiryDate")  & ""	
	sCSC = adoData( "CSC")  & ""
	sCardAddress = adoData( "CardAddress")  & ""
	sCardPostCode = adoData( "CardPostCode")  & ""
	sCardType = adoData( "CardType")  & ""

	
	'sOR
	call SetObjRepository ("PCI",sProductDir & "Resources\")

	If Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").Exist(5) Then
		If sEnv = "E4" Then
			Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebEdit("CardNumber").WebSet sCardNumberE4
		ElseIf sEnv = "E7" Then
			Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebEdit("CardNumber").WebSet sCardNumberE7
		ElseIf sEnv = "C2" Then
			Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebEdit("CardNumber").WebSet sCardNumberC2
		End If
		Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebEdit("ExpiryDate").WebSet sExpiryDate
		Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebEdit("CSC").WebSet sCSC
		Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebEdit("CardAddress").WebSet sCardAddress
		Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebEdit("CardPostCode").WebSet sCardPostCode
		CaptureSnapshot()
		AddVerificationInfoToResult  "Info" ,"Details are entered successfully."
		Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:").WebButton("Submit").Click
	Else
		AddVerificationInfoToResult  "Error" ,"Card Details button is not clicked successfully."
		iModuleFail = 1
		Exit Function
	End If

	If Browser("Aeriandi Secure Call Control:").Page("Aeriandi Secure Call Control:_2").WebElement("success").Exist(5) Then
		AddVerificationInfoToResult  "Info" ,"CardDetails are entered successfully."
	Else
		AddVerificationInfoToResult  "Error" ,"CardDetails are not  entered successfully."
		iModuleFail = 1
	End If

	CaptureSnapshot()

	Browser("Aeriandi Secure Call Control:").Close


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

	

End Function


'#################################################################################################
' 	Function Name : AuthorizeCreditCard_fn
' 	Description : This function is used to click on Authorize button and capture the card details
'   Created By :  Tarun Bansal
'	Creation Date :        
'##################################################################################################
Public Function AuthorizeCreditCard_fn()
   	'Get Data
	Dim adoData	  
'    strSQL = "SELECT * FROM [Aeriandi$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\PCI.xls",strSQL)

	call SetObjRepository ("PCI",sProductDir & "Resources\")

	If SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Credit/Debit Card Payment").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Credit/Debit Card Payment").SiebButton("Authorize").SiebClick False
	Else
		AddVerificationInfoToResult  "Error" , "Authorize button is not present on page."
		iModuleFail = 1
		Exit Function
	End If

	sCardNumber = SiebApplication("Siebel Call Center").SiebScreen("Accounts").SiebView("Account Billing Profile").SiebApplet("Credit/Debit Card Payment").SiebText("Card no.").GetRoProperty("text")
	If sCardNumber <> "" Then
		AddVerificationInfoToResult  "Info" , "Card Number : - " & sCardNumber & " is  updated successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Authorize button is not clicked successfully."
		iModuleFail = 1
	End If

	CaptureSnapshot()


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

	

End Function



'#################################################################################################
' 	Function Name : AuthorizePaymentDetailsCreditCard_fn
' 	Description : This function is used to click on Authorize button in payments tab and capture the card details
'   Created By :  Tarun Bansal
'	Creation Date :        
'##################################################################################################
Public Function AuthorizePaymentDetailsCreditCard_fn()
   	'Get Data
	Dim adoData	  
'    strSQL = "SELECT * FROM [Aeriandi$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\PCI.xls",strSQL)

	call SetObjRepository ("Orders",sProductDir & "Resources\")

	If SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Details - Credit").Exist(5) Then
		SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Details - Credit").SiebButton("Authorise").SiebClick False
	Else
		AddVerificationInfoToResult  "Error" , "Authorize button is not present on page."
		iModuleFail = 1
		Exit Function
	End If


	sCardNumber = SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Payment").SiebApplet("Payment Details - Credit").SiebText("Card no.").GetRoProperty("text")
	If sCardNumber <> "" Then
		AddVerificationInfoToResult  "Info" , "Card Number : - " & sCardNumber & " is  updated successfully."
	Else
		AddVerificationInfoToResult  "Error" , "Authorize button is not clicked successfully."
		iModuleFail = 1
	End If

	CaptureSnapshot()


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

	

End Function