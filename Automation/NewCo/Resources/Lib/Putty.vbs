'#################################################################################################
' 	Function Name : LoginToPutty_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function LoginToPutty_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [LoginToPutty$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
	
	'sOR
	
	call SetObjRepository ("Putty",sProductDir & "Resources\")

	sHostName = adoData("HostName"&sEnv) & ""
	sPort = adoData("Port") & ""
	sLogFileName = adoData("LogFileName") & ""
	sLogPath = Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\" & sLogFileName & ".log"
	sUsername = sBRMPuttyLogin
	sPassphrase = sBRMPuttyPassphrase
	sNumber = adoData("Number") & ""
	sRootUser = adoData("RootUser") & ""
	sRootPassword = adoData("RootPassword") & ""
	Call Kill_OpenPuttyInstances()

	SystemUtil.Run "C:\Program Files (x86)\PuTTY\putty.exe"
	Window("PuTTY Configuration").Activate
	
	Window("PuTTY Configuration").WinEdit("Host Name (or IP address)").WinSet sHostName
	Window("PuTTY Configuration").WinEdit("Port").WinSet sPort
	Window("PuTTY Configuration").WinTreeView("Category:").WinSelect "Session;Logging"
	Window("PuTTY Configuration").WinRadioButton("All session output").WinRadioSet
	Window("PuTTY Configuration").WinEdit("Log file name:").WinSet sLogPath
	Window("PuTTY Configuration").WinRadioButton("Always append to the end").WinRadioSet
	Window("PuTTY Configuration").WinTreeView("Category:").Expand "Connection;SSH"
	Window("PuTTY Configuration").WinTreeView("Category:").WinSelect "Connection;SSH;Auth"

	Window("PuTTY Configuration").WinCheckBox("Allow agent forwarding").Click
	Window("PuTTY Configuration").WinCheckBox("Allow attempted changes").Click

	Window("PuTTY Configuration").WinEdit("Private key file for authentic").WinSet "C:\Users\" &sUsername & "\PuTTY\ssh_key_putty_newVo" & sEnv & ".ppk"
	
	Window("PuTTY Configuration").WinButton("Open").WinClick
'	Set oFSOPuttyLog = CreateObject("Scripting.FileSystemObject")
'	Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
	wait 2
	'Window("PuTTY").PuttyType "login as:",sUsername
	'Window("PuTTY").PuttyType "password:",sPassphrase

	Window("PuTTY").PuttyType "login as:","v.a.ranjan"
	Window("PuTTY").PuttyType "password:","P@ssw0rd"

	Window("PuTTY").Maximize
'	Window("PuTTY").PuttyType  "Exist:Passphrase for key",sPassphrase
	REM If Ucase(sEnv)<>"SUP02" Then
		REM Window("PuTTY").PuttyTypeOnly  "enter a number",sNumber
		REM Window("PuTTY").PuttyTypeOnly  "Connect to","Y"
	REM End If
	
	'Window("PuTTY").PuttyType  "password:",sRootUser
	Window("PuTTY").PuttyType  "$",sRootUser
	Window("PuTTY").PuttyType  "password:",sRootPassword
'	CreateLogFile("Login")
	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : Rating_fn1
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
'Public Function Rating_fn1()
'
''Get Data
'	Dim adoData	  
'    strSQL = "SELECT * FROM [Rating$] WHERE  [RowNo]=" & iRowNo
'	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
'
'	sStreamChoice = adoData("StreamChoice") & ""
'	sCDRType = adoData("CDRType") & ""
'	sNoEDR = adoData("NoEDR") & ""
'	sDomesticRoaming = adoData("Domestic/Roaming") & ""
'
'	sCountryCode = adoData("CountryCode") & ""
'	
'
''	DictionaryTest_G("ACCNTMSISDN")="447387960383"
''	DictionaryTest_G("AccountNo") = "7000226867"
'
'	sCallingNo = adoData("CallingNo") & ""
'	If DictionaryTest_G.Exists(sCallingNo) Then
'		sCallingNo = Replace (sCallingNo,sCallingNo,DictionaryTest_G(sCallingNo))
'	End If
'
'	sCalledNo = adoData("CalledNo") & ""
'	If DictionaryTest_G.Exists(sCalledNo) Then
'		sCalledNo = Replace (sCalledNo,sCalledNo,DictionaryTest_G(sCalledNo))
'	End If
'
'	sCnumber = adoData("Cnumber") & ""
'	If DictionaryTest_G.Exists(sCnumber) Then
'		sCnumber = Replace (sCnumber,sCnumber,DictionaryTest_G(sCnumber))
'	End If
'
'	sSMSQuantity = adoData("SMSQuantity") & ""
'	sGSM = adoData("GSM") & ""
'	sDuration = adoData("Duration") & ""
'	sCDRType = adoData("CDRType") & ""	
'	sNoEDR = adoData("NoEDR") & ""		
'	sTariffClass = adoData("TariffClass") & ""
'	sSrcCountry = adoData("SrcCountry") & ""
'	sDetnCountry = adoData("DetnCountry") & ""
''	sKey = adoData("CDRKey") & ""
'	sRateGBP = adoData("RateGBP") & ""
'	iIndex = adoData( "Index")  & ""
'    sInOutPartial = adoData("InOutPartial") & ""
'	sTransactionType =  adoData("TransactionType") & ""
'
'	If iIndex="" Then
'		iIndex=0
''		else
''			iIndex=Cint(iIndex)
'	End If
'	'sOR
'	call SetObjRepository ("Putty",sProductDir & "Resources\")
'	Window("PuTTY").Activate
'	Window("PuTTY").PuttyType  "","date +%s"
'	CapturePuttyData "LINEAFTER:date +%s","Timestamp",iIndex
'	DictionaryTest_G("Timestamp") = DictionaryTest_G("Timestamp")+86400
'	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/rating"
'    Window("PuTTY").PuttyType  "$","./rating.sh"
'	Window("PuTTY").PuttyType  "",sStreamChoice
'	Window("PuTTY").PuttyType  "",sCDRType
'	Window("PuTTY").PuttyType  "EDR",sNoEDR
'	Window("PuTTY").PuttyType  "Domestic/Roaming",sDomesticRoaming
'	Window("PuTTY").PuttyType  "code",sCountryCode
'	Window("PuTTY").PuttyType  "In/Out/Partial",sInOutPartial
'	Window("PuTTY").PuttyType  "A Number",sCallingNo
'	Window("PuTTY").PuttyType  "B Number",sCalledNo
'	Window("PuTTY").PuttyType  "C Number",sCnumber
'	Window("PuTTY").PuttyType  "Rate in GBP",sRateGBP
'	Window("PuTTY").PuttyType  "Transaction Type",sTransactionType
'	Window("PuTTY").PuttyType  "quantity",sSMSQuantity
'	Window("PuTTY").PuttyType  "Duration",sDuration
'	Window("PuTTY").PuttyType  "Timestamp",DictionaryTest_G("Timestamp")
'	Window("PuTTY").PuttyType  "GSM",sGSM
'	Window("PuTTY").PuttyType  "Tariff class",sTariffClass
'	Window("PuTTY").PuttyType  "Source country",sSrcCountry
'	Window("PuTTY").PuttyType  "Destination country",sDetnCountry
'
'	
'		
'
'	If Err.Number <> 0 Then
'            iModuleFail = 1
'			AddVerificationInfoToResult  "Error" , Err.Description
'	End If
'	
'End Function
'


'#################################################################################################
' 	Function Name : Rating_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function Rating_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Rating$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)

	sStreamChoice = adoData("StreamChoice") & ""
	sCDRType = adoData("CDRType") & ""
	sNoEDR = adoData("NoEDR") & ""
	sDomesticRoaming = adoData("Domestic/Roaming") & ""
	sCr_Dr = adoData("CrDr") & ""
	sTax = adoData("Tax") & ""
	sDestination = adoData("Destination") & ""
	sBearerCapability = adoData("BearerCapability") & ""

	sCountryCode = adoData("CountryCode") & ""
	If sCountryCode ="" Then
		sCountryCode = DictionaryTest_G("CountryCode")
	End If	

'	DictionaryTest_G("ACCNTMSISDN")="447387960912"
'	DictionaryTest_G("AccountNo") = "7000228866"

	sCallingNo = adoData("CallingNo") & ""
	If DictionaryTest_G.Exists(sCallingNo) Then
		sCallingNo = Replace (sCallingNo,sCallingNo,DictionaryTest_G(sCallingNo))
	End If

	DictionaryTest_G ("sCallingNo") = sCallingNo

	sCalledNo = adoData("CalledNo") & ""
	If DictionaryTest_G.Exists(sCalledNo) Then
		sCalledNo = Replace (sCalledNo,sCalledNo,DictionaryTest_G(sCalledNo))
	End If

	sTimestamp = adoData("Timestamp") & ""

	sCnumber = adoData("Cnumber") & ""
	If DictionaryTest_G.Exists(sCnumber) Then
		sCnumber = Replace (sCnumber,sCnumber,DictionaryTest_G(sCnumber))
	End If

	sSMSQuantity = adoData("SMSQuantity") & ""
	If sSMSQuantity<>"" Then
		DictionaryTest_G("Duration_Quantity") = sSMSQuantity
	End If

	sGSM = adoData("GSM") & ""
	If sGSM="" Then
		sGSM = DictionaryTest_G("GSM")
	End If

	sDuration = adoData("Duration") & ""
	If sDuration<>"" Then
		DictionaryTest_G("Duration_Quantity") = sDuration
	End If

	sCDRType = adoData("CDRType") & ""	
	sNoEDR = adoData("NoEDR") & ""		
	sTariffClass = adoData("TariffClass") & ""
	If sTariffClass ="" Then
		sTariffClass = DictionaryTest_G("TariffClass")
	End If
	
	sSrcCountry = adoData("SrcCountry") & ""
'	sSrcCountry = DictionaryTest_G("SrcCountry")
	If sSrcCountry ="" Then
		sSrcCountry = DictionaryTest_G("SrcCountry")
	End If
	sDetnCountry = adoData("DetnCountry") & ""
	If sDetnCountry="" Then
		sDetnCountry = DictionaryTest_G("DestnCountry")
	End If
	
'	sKey = adoData("CDRKey") & ""
	sRateGBP = adoData("RateGBP") & ""
	DictionaryTest_G("RateGBP") = sRateGBP
	iIndex = adoData( "Index")  & ""
    sInOutPartial = adoData("InOutPartial") & ""
	sTransactionType =  adoData("TransactionType") & ""

	If iIndex="" Then
		iIndex=0
'		else
'			iIndex=Cint(iIndex)
	End If
	'sOR
	call SetObjRepository ("Putty",sProductDir & "Resources\")

	Window("PuTTY").Activate

	If sTimestamp = "" Then
		
		Window("PuTTY").PuttyType  "$","date +%s"
		wait 2
		Window("PuTTY").PuttyType  "$","abcd"
		'Window("PuTTY").PuttyType  "","echo " & chr(34) & "Your Timestamp is " & chr(34) & chr(96) & "date +%s" & chr(96)
		'CapturePuttyData = CapturePuttyData ("Your Timestamp is ","TIMESTAMP0",iIndex)
		Call AccntCreationDate_fn()
		CapturePuttyData "LINEAFTER:date +%s","TIMESTAMP0",iIndex
		If  DictionaryTest_G("CreationDate") > DictionaryTest_G("TIMESTAMP0") Then
			DictionaryTest_G("TIMESTAMP0") = DictionaryTest_G("TIMESTAMP0") + 86400
		else
'			DictionaryTest_G("TIMESTAMP0") = DictionaryTest_G("CreationDate")
				DictionaryTest_G("TIMESTAMP0") = DictionaryTest_G("TIMESTAMP0") + 14400
		End If

		If  DictionaryTest_G("CreationDate")="" Then
			DictionaryTest_G("TIMESTAMP0") = DictionaryTest_G("TIMESTAMP0") + 86400
		End If

	End If
	
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/rating"
    Window("PuTTY").PuttyType  "$","./rating.sh"
	Window("PuTTY").PuttyType  "",sStreamChoice
	Window("PuTTY").PuttyType  "",sCDRType
	Window("PuTTY").PuttyType  "EDR",sNoEDR
	Window("PuTTY").PuttyType  "Domestic/Roaming",sDomesticRoaming
	Window("PuTTY").PuttyType  "code",sCountryCode
	Window("PuTTY").PuttyType  "In/Out/Partial",sInOutPartial
	Window("PuTTY").PuttyType  "A Number",sCallingNo
	Window("PuTTY").PuttyType  "B Number",sCalledNo
	Window("PuTTY").PuttyType  "C Number",sCnumber
	Window("PuTTY").PuttyType  "credit/debit",sCr_Dr
	Window("PuTTY").PuttyType  "Duration",sDuration
	Window("PuTTY").PuttyType  "quantity",sSMSQuantity
	Window("PuTTY").PuttyType  "Rate in GBP",sRateGBP
	Window("PuTTY").PuttyType  "Transaction Type",sTransactionType
'	Window("PuTTY").PuttyType  "quantity",sSMSQuantity
	Window("PuTTY").PuttyType  "Timestamp",DictionaryTest_G("TIMESTAMP0")
	Window("PuTTY").PuttyType  "Destination",sDestination
	Window("PuTTY").PuttyType  "Bearer",sBearerCapability
	Window("PuTTY").PuttyType  "tax",sTax
	Window("PuTTY").PuttyType  "GSM",sGSM
	Window("PuTTY").PuttyType  "Tariff class",sTariffClass
	Window("PuTTY").PuttyType  "Source country",sSrcCountry
	Window("PuTTY").PuttyType  "Destination country",sDetnCountry

	
	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function

'#################################################################################################
' 	Function Name : CheckRatingOutputFile_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CheckRatingOutputFile_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Rating$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
	iIndex = adoData( "Index")  & ""
	If iIndex="" Then
		iIndex=0
'		else
'			iIndex=Cint(iIndex)
	End If
'	sKey = adoData("CDRKey") & ""
'	sCDRSeqKey = adoData("CDRSeqKey") & ""

'	CapturePuttyData ".data","CDR",iIndex
	sCapturePuttyData = CapturePuttyData ("LINEAFTER:Your CDR is","CDR",iIndex)
	If Not(sCapturePuttyData = True) Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "CDR is not present"
		Exit Function
	End If
	sCDR = DictionaryTest_G("CDR" )
	call SetObjRepository ("Putty",sProductDir & "Resources\")

	Window("PuTTY").Activate

	Window("PuTTY").PuttyType "$","cd /home/pin/opt/portal/var/brm/acrinput/router"
	
	Window("PuTTY").PuttyType "$","cat " & sCDR
	
	'Window("PuTTY").PuttyType "$","vi " & sCDR
	'Window("PuTTY").PuttyType "",micEsc
	wait 1
	CaptureSnapshot()
	'Window("PuTTY").PuttyType "",":q!"
	

	sCDR = Replace (sCDR,".data","")
	sCDR = Replace (sCDR,Split(sCDR,"_")(0),"")
	
	DictionaryTest_G("CDRSeq") = CLng(Split(sCDR,"_")(Ubound(Split(sCDR,"_"))))
 

	Window("PuTTY").PuttyType "$","cd /home/pin/opt/portal/var/brm/ploutput/pin01/pin_rel_01/archive"
'	Window("PuTTY").PuttyType  "$","ls -ltr *" & sCDR & "*"

	Set obj = Window("PuTTY")
	sTypeCapturePuttyData = TypeCapturePuttyData (obj,"$","ls -ltr *" & sCDR & "*",sCDR &".out.bc","OutCDR")
	If Not(sTypeCapturePuttyData = True) Then
'		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "CDR is not present"
'		Exit Function
	End If
	outCDR = DictionaryTest_G("OutCDR")
	If outCDR<>"" Then

		outCDR = Replace(Split(outCDR," ")(Ubound(Split(outCDR," "))),"[00m","")
		Window("PuTTY").PuttyType  "$","cat " & outCDR	
		'Window("PuTTY").PuttyType  "$","vi " & outCDR	
		
		'Window("PuTTY").PuttyType "",micEsc
		wait 1
		CaptureSnapshot()
		'Window("PuTTY").PuttyType "",":q!"
	End If

'	CreateLogFile("Rating")
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : Billing_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function Billing_fn()

'Get Data
 
 
	'sOR
	Dim adoData	  
    strSQL = "SELECT * FROM [Billing$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
	sAccnt = adoData( "Accnt")  & ""
	If DictionaryTest_G.Exists(sAccnt) Then
		sAccnt = Replace (sAccnt,sAccnt,DictionaryTest_G(sAccnt))
	End If
	If DictionaryTempTest_G.Exists(sAccnt) Then
		sAccnt = Replace (sAccnt,sAccnt,DictionaryTempTest_G(sAccnt))
	End If
	iIndex = adoData( "Index")  & ""
	If iIndex="" Then
		iIndex=0
'		else
'			iIndex=Cint(iIndex)
	End If

	sSequence = adoData( "Sequence")  & ""
	If sSequence="" Then
		sSequence=1
	End If

	call SetObjRepository ("Putty",sProductDir & "Resources\")

	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/bill_invoice"
    Window("PuTTY").PuttyType  "bill_invoice","./billing.sh"
	Window("PuTTY").PuttyType  "Account",sAccnt
	Window("PuTTY").PuttyType  "sequence",sSequence

'	Window("PuTTY").PuttyType  "Account","7000226156"

	sValidatePutty = ValidatePutty ("","Total number of records processed = 1")
	If Not(sValidatePutty = True) Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Billing Failed"
	End If

	sCapturePuttyData = CapturePuttyData ("LINEAFTER:Bill_No","BillNo",iIndex)
	CaptureSnapshot()

	If Not(sCapturePuttyData = True) Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Billing Failed"
	End If

	DictionaryTest_G.Item("BillNo")=BillNo
	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : Invoicing_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function Invoicing_fn()

'Get Data
 	Dim adoData	  
    strSQL = "SELECT * FROM [Invoicing$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
	sAccnt = adoData( "Accnt")  & ""
	If DictionaryTest_G.Exists(sAccnt) Then
		sAccnt = Replace (sAccnt,sAccnt,DictionaryTest_G(sAccnt))
	elseif DictionaryTempTest_G.Exists(sAccnt) Then
		sAccnt = Replace (sAccnt,sAccnt,DictionaryTempTest_G(sAccnt))
	End If

	iIndex = adoData( "Index")  & ""
	If iIndex="" Then
		iIndex=0
	End If

	sSequence = adoData( "Sequence")  & ""
	If sSequence="" Then
		sSequence=1
	End If

	sInvoiceTimeSeq = adoData( "InvoiceTimeSeq")  & ""
	If sInvoiceTimeSeq="" Then
		sInvoiceTimeSeq="1;0"
	End If
 
	'sOR
	call SetObjRepository ("Putty",sProductDir & "Resources\")

	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/bill_invoice"
    Window("PuTTY").PuttyType  "bill_invoice","./invoicing.sh"
	Window("PuTTY").PuttyType  "Account",sAccnt
	Window("PuTTY").PuttyType  "sequence of bill no for invoicing",sSequence
	Window("PuTTY").PuttyType  "invoicing time (seq;sec)",sInvoiceTimeSeq
'	Window("PuTTY").PuttyType  "Account","7000226156"

	sValidatePutty = ValidatePutty ("Account no for Invoicing","Total number of records processed = 1")
	If Not(sValidatePutty = True) Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Invoicing Failed"
	End If
	sCapturePuttyData = CapturePuttyData ("Invoice poid is","PoID",iIndex)
	DictionaryTest_G("PoID") = Replace (DictionaryTest_G("PoID"),"Invoice poid is ","")
	CaptureSnapshot()

	If Not(sCapturePuttyData = True) Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Invoicing Failed"
	End If
	
	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : test_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function test_fn()

	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").PuttyType  "","date +%s"
	CapturePuttyData "LINEAFTER:date +%s","Timestamp",3
	print DictionaryTest_G("Timestamp")
	Window("PuTTY").PuttyType  "","date +%s"
	CapturePuttyData "LINEAFTER:date +%s","Timestamp",9
	print DictionaryTest_G("Timestamp")
	CapturePuttyData "LINEAFTER:date +%s","Timestamp",10
	print DictionaryTest_G("Timestamp")
	CapturePuttyData "LINEAFTER:date +%s","Timestamp",11
	print DictionaryTest_G("Timestamp")
	CapturePuttyData "LINEAFTER:date +%s","Timestamp",12
	print DictionaryTest_G("Timestamp")

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function




'#################################################################################################
' 	Function Name : Purchased_Product_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function Purchased_Product_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate

'	DictionaryTest_G.Item("AccountNo")="7000257292"
'	DictionaryTest_G.Item("ACCNTMSISDN")="447389703066"
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/sql_query"
	Window("PuTTY").PuttyType  "$","./purchased_product.sh " & DictionaryTest_G.Item("AccountNo")
	wait 2
	CaptureSnapshot()
End Function





'#################################################################################################
' 	Function Name : ExitPutty_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function ExitPutty_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","exit"
	Window("PuTTY").PuttyType  "$","exit"
	Window("PuTTY").PuttyTypeOnly  "enter a number",99
	Window("PuTTY").PuttyTypeOnly  "Exit menu","y"
End Function




'#################################################################################################
' 	Function Name : BillNow_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function BillNow_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/bill_invoice"
	Window("PuTTY").PuttyType  "$","./bill_now.sh"

	sValidatePutty = ValidatePutty ("bill_now.sh","Bill now done")
	CaptureSnapshot()
	If sValidatePutty = True Then
		AddVerificationInfoToResult  "Info" , "Bill Now done"
	Else
		AddVerificationInfoToResult  "Error" , "Bill now not completed"
		iModuleFail = 1
		Exit Function
	End If


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function



'#################################################################################################
' 	Function Name : AdjustmentDetails_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function AdjustmentDetails_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/sql_query"


	'DictionaryTest_G("AccountNo") = "7000219784"
	'DictionaryTest_G("ReasonCode")="200094"
	'DictionaryTest_G("Reason")= "UK Data, Content, MMS Credit" 
	sAccnt = DictionaryTest_G.Item("AccountNo")
	Window("PuTTY").PuttyType  "$","./adjustment_detail.sh " & sAccnt & " " & DictionaryTest_G.Item("ReasonCode") & " " & Chr(34)  &  DictionaryTest_G.Item("Reason")  & Chr(34)

	sValidatePutty = ValidatePutty ("adjustment_detail.sh","Adjustment detail for account " & sAccnt)
	CaptureSnapshot()
	If sValidatePutty = True Then
		AddVerificationInfoToResult  "Info" , "Adjustment done in BRM"
	Else
		AddVerificationInfoToResult  "Error" , "Adjustment not done in BRM"
		iModuleFail = 1
		Exit Function
	End If


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function




'#################################################################################################
' 	Function Name : OverPaymentAllocation_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function OverPaymentAllocation_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/appsCustom/vf_bulk_overpayment_allocation"
	Window("PuTTY").PuttyType  "$","./vf_bulk_overpayment_allocation -verbose"
	
	sValidatePutty = ValidatePutty ("vf_bulk_overpayment_allocation -verbose","Total number of records processed")
	CaptureSnapshot()
	If sValidatePutty = True Then
		AddVerificationInfoToResult  "Info" , "OverPayment Allocation executed successfully"
	Else
		AddVerificationInfoToResult  "Error" , "OverPayment Allocation not executed successfully"
		iModuleFail = 1
		Exit Function
	End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : AccountChargeApply_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function AccountChargeApply_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd $PIN_HOME/apps/pin_billd"
	Window("PuTTY").PuttyType  "$","pin_cycle_fees -purchase -verbose"

	sValidatePutty = ValidatePutty ("pin_cycle_fees -purchase -verbose","Total number of records processed")
	CaptureSnapshot()
	If sValidatePutty = True Then
		AddVerificationInfoToResult  "Info" , "pin_cycle_fees executed successfully"
	Else
		AddVerificationInfoToResult  "Error" , "pin_cycle_fees not executed successfully"
		iModuleFail = 1
		Exit Function
	End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : Payment_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################

Public Function Payment_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Payment$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)

	sPaymentType = adoData("PaymentType") & ""
	sAccountNo = adoData("AccountNo") & ""
	sAmount = adoData("Amount") & ""
	sBillNoOption = adoData("BillNoOption") & ""

'	If sBillNoOption="" Then
'		sBillNoOption="N"
'	End If

	If DictionaryTest_G.Exists("OverPayment") Then
		DictionaryTest_G.Item("OverPayment")=sAmount
	else
	DictionaryTest_G.add "OverPayment",sAmount
	End If	

	If DictionaryTest_G.Exists(sAmount) Then
		sAmount = Replace (sAmount,sAmount,DictionaryTest_G(sAmount))
	else
			If sAmount="" Then
				'sAmount=DictionaryTest_G ("AMOUNT0")
				sAmount=""
			Else
				DictionaryTest_G("AMOUNT0")=sAmount
			End If	
	End If

		
	

'	DictionaryTest_G("ACCNTMSISDN")="447387960383"
'	DictionaryTest_G("AccountNo") = "7000226867"

	If sAccountNo = "ACCNTMSISDN" Then
		sAccountNo = DictionaryTest_G ("AccountNo")
'	
	End If

       	
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/payment"
    Window("PuTTY").PuttyType  "$","./payments.sh"
	Window("PuTTY").PuttyType  "",sPaymentType
	CaptureSnapshot()
	Window("PuTTY").PuttyType  "Account no",sAccountNo
	Window("PuTTY").PuttyType  "(Y/N)",sBillNoOption
	Window("PuTTY").PuttyType  "amount in GBP",sAmount	

	sValidatePutty = ValidatePutty ("./payments.sh ","payment done")
	CaptureSnapshot()
		If sValidatePutty = True Then
			AddVerificationInfoToResult  "Info" , "Payment done for your account  " & sAccountNo
		else
			AddVerificationInfoToResult  "Error" , "Payment not done for your account  " & sAccountNo
			iModuleFail = 1
			Exit Function
		End If



	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function






'#################################################################################################
' 	Function Name : SBExportReport_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################

Public Function SBExportReport_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [SBExportReport$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)

	sReportOption  = adoData("ReportOption") & ""
	sReportType = adoData("ReportType") & ""
	sAllSegments = adoData("AllSegments") & ""
	iIndex = adoData( "Index")  & ""

'	DictionaryTest_G("ACCNTMSISDN")="447387960383"
'	DictionaryTest_G("AccountNo") = "7000226867"

	sAccountNo = DictionaryTest_G ("AccountNo")

       	
	call SetObjRepository ("Putty",sProductDir & "Resources\")

	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/edw"
'	If  sReportOption = 1 or sReportOption = 3  Then
		Window("PuTTY").PuttyType  "$","./SBExport_report.sh " & DictionaryTest_G.Item("AccountNo")
		Window("PuTTY").PuttyType  "",sReportOption
		Window("PuTTY").PuttyType  "",sReportType
		If sReportOption = "4" Then
			If sAllSegments = "" Then
				sAllSegments = "N"
			End If
			wait 2
			Window("PuTTY").PuttyType  "",sAllSegments
		End If

'	elseIf sReportOption = 2 Then
'		Window("PuTTY").PuttyType  "$","./SBExport_report.sh " & DictionaryTest_G.Item("AccountNo")
'		Window("PuTTY").PuttyType  "",sReportOption	
'	End If

	sValidatePutty = ValidatePutty ("SBExport.sh","SBExport file is")

'	If  sReportOption = 1 or sReportOption = 2 Then
		If sValidatePutty = True Then
			AddVerificationInfoToResult  "Info" , "SBExport report generated successfully for your account " & sAccountNo
		else
			AddVerificationInfoToResult  "Error" , "SBExport report could not generate for your account  " & sAccountNo
			iModuleFail = 1
			Exit Function
		End If
'	elseIf sReportOption = 3 Then
'			If sValidatePutty = True Then
'				AddVerificationInfoToResult  "Info" , "SBExport report generated successfully "
'			else
'				AddVerificationInfoToResult  "Error" , "SBExport report could not generate"
'				iModuleFail = 1
'				Exit Function
'			End If
'	End If
	
	sCapturePuttyData = CapturePuttyData ("SBExport file is","ReportFile",iIndex)

	DictionaryTest_G ("ReportFile") = Replace(DictionaryTest_G ("ReportFile"),"SBExport file is ","")

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function



'#################################################################################################
' 	Function Name : UnixTimestamp_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################

Public Function UnixTimestamp_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Timestamp$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)


	iDays  = adoData("Days") & ""
	iHr  = adoData("Hr") & ""
	iMin  = adoData("Min") & ""
	iSec  = adoData("Sec") & ""
	sWeekDay  = adoData("WeekDay") & ""

	DictionaryTest_G ("Days") = iDays
	DictionaryTest_G ("Hr") = iHr
	DictionaryTest_G ("Min") = iMin
	DictionaryTest_G ("Sec") = iSec
	DictionaryTest_G ("WeekDay") = sWeekDay

End Function



'
'#################################################################################################
' 	Function Name : CreditAlert_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function CreditAlert_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	


	Dim adoData	  
    strSQL = "SELECT * FROM [CreditAlert$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)

	sOption = adoData("Option") & ""
	sAccountNo = DictionaryTest_G ("AccountNo")
	sAmount = adoData("Amount") & ""
	sThreshold = adoData("Threshold") & ""
	sBusinessType = adoData("BusinessType") & ""
	sCRIValue = adoData("CRIValue") & ""
	sMSISDN = adoData("MSISDN") & ""
	If sMSISDN = "MSISDN0" Then
		sMSISDN = DictionaryTest_G("ACCNTMSISDN0")
	else
		sMSISDN = DictionaryTest_G("ACCNTMSISDN")
	End If
'	sMSISDN = DictionaryTest_G("ACCNTMSISDN")

'     If DictionaryTest_G.Exists(sMSISDN) Then
'		sMSISDN = Replace (sMSISDN,sMSISDN,DictionaryTest_G(sMSISDN))
'	End If  


	'DictionaryTest_G("AccountNo") = "7000219784"
	'DictionaryTest_G("ACCNTMSISDN")="447387960383"
	'sAccountNo = DictionaryTest_G ("AccountNo")
	'sMSISDN = DictionaryTest_G("ACCNTMSISDN")

	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/credit_alert"
	
	Window("PuTTY").PuttyType  "$","./credit_alert.sh " 
	Window("PuTTY").PuttyType  "",sOption
	If  sOption = 1 Then
		Window("PuTTY").PuttyType  "Account Number",sAccountNo
		Window("PuTTY").PuttyType  "CRI Value",sCRIValue
	elseif sOption = 2 Then
		Window("PuTTY").PuttyType  "MSISDN",sMSISDN
		Window("PuTTY").PuttyType  "AMOUNT",sAmount
		Window("PuTTY").PuttyType  "THRESHOLDS",sThreshold
		Window("PuTTY").PuttyType  "BUSINESSTYPE",sBusinessType
	End If

	
	If sOption = 1  Then
		sValidatePutty = ValidatePutty ("credit_alert.sh","CRI Value changed to " & sCRIValue)
		CaptureSnapshot()
		If sValidatePutty = True Then
			AddVerificationInfoToResult  "Info" , "CRI Value changed successfully"
		else
			AddVerificationInfoToResult  "Error" , "CRI value could not changed"
			iModuleFail = 1
			Exit Function
		End If
	Elseif sOption = 2 Then
		sValidatePutty = ValidatePutty ("credit_alert.sh","Credit limit changed to " & sAmount)
		CaptureSnapshot()
		If sValidatePutty = True Then
			AddVerificationInfoToResult  "Info" , "Credit limit changed successfull"
		else
			AddVerificationInfoToResult  "Error" , "Credit Limit could not changed"
			iModuleFail = 1
			Exit Function
		End If

	End If

		

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function







'#################################################################################################
' 	Function Name : CreditAlertSiebel_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function CreditAlertSiebel_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	

	'DictionaryTest_G("AccountNo") = "7000219784"
	'DictionaryTest_G("ACCNTMSISDN")="447387946254"
	'sAccountNo = DictionaryTest_G ("AccountNo")

	sMSISDN = DictionaryTest_G("ACCNTMSISDN")



	Window("PuTTY").PuttyType  "$","cd /opt/crm/siebel/batchfs/automation"
	
	Window("PuTTY").PuttyType  "$","./ca.sh "  & sMSISDN
    
	sValidatePutty = ValidatePutty ("./ca.sh " & sMSISDN,"Alert generated for your MSISDN " & sMSISDN)
		CaptureSnapshot()
		If sValidatePutty = True Then
			AddVerificationInfoToResult  "Info" , "Alert generated for your MSISDN  " & sMSISDN
		else
			AddVerificationInfoToResult  "Error" , "Alert not generated for your MSISDN  " & sMSISDN
			iModuleFail = 1
			Exit Function
		End If


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function


'#################################################################################################
' 	Function Name : EDWUsageFile_fn
' 	Description : 
'   Created By :  Tarun
'	Creation Date :        
'##################################################################################################
Public Function EDWUsageFile_fn()


	call SetObjRepository ("Putty",sProductDir & "Resources\")

	sCallingNo = DictionaryTest_G ("sCallingNo")

	iIndex = 0

	Window("PuTTY").Activate

	
	Window("PuTTY").PuttyType  "$","cd $PIPE_DIR/ploutput/aggregation/out/edw"
    Window("PuTTY").PuttyType  "$","./EDW.sh " & sCallingNo
	
	sCapturePuttyData = CapturePuttyData ("Your File Name","ReportFile",iIndex)

	DictionaryTest_G ("ReportFile") = TRIM(Replace(DictionaryTest_G ("ReportFile"),"Your File Name ",""))

	If Trim(DictionaryTest_G ("ReportFile")) = "" Then
		AddVerificationInfoToResult  "Error" , "File is not generated successfully."
		iModuleFail = 1
	End If


	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function


'#################################################################################################
' 	Function Name : GLReport_fn
' 	Description : 
'   Created By :  Vivek Ranjan
'	Creation Date :        
'##################################################################################################
Public Function GLReport_fn()


	call SetObjRepository ("Putty",sProductDir & "Resources\")



	Dim adoData	  
    strSQL = "SELECT * FROM [GL_Report$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
	sBusinessProfileId = adoData( "BusinessProfileId")  & ""


	iIndex = 0

	Window("PuTTY").Activate

	
	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/edw"
    Window("PuTTY").PuttyType  "$","./GL_Report.sh"
	Window("PuTTY").PuttyType  "",sBusinessProfileId
	sCapturePuttyData = CapturePuttyData ("GL Report Zip File is ","ReportFile",iIndex)

	DictionaryTest_G ("ReportFile") = TRIM(Replace(DictionaryTest_G ("ReportFile"),"GL Report Zip File is ",""))


	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
	
End Function



'#################################################################################################
' 	Function Name : Docgen_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function Docgen_fn()
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	

'	Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Docgen$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)

	sAccountNo  = adoData("Account") & ""
	sSequence = adoData("Sequence") & ""
	iIndex = adoData( "Index")  & ""
	If iIndex="" Then
		iIndex=0
	End If

'	DictionaryTest_G("AccountNo") = "7000226867"

	'If iIndex="AccountNo" Then
		sAccountNo = DictionaryTest_G ("AccountNo")
'	End If	
	

       	
	call SetObjRepository ("Putty",sProductDir & "Resources\")

	Window("PuTTY").PuttyType  "$","cd /home/pin/opt/portal/7.5/automation/bill_invoice"
		Window("PuTTY").PuttyType  "$","./docgen.sh"
		Window("PuTTY").PuttyType  "enter Account",sAccountNo
		Window("PuTTY").PuttyType  "enter sequence",sSequence


	sValidatePutty = ValidatePutty ("docgen.sh","Docgen completed successfully please check in WCC",iIndex)


		If sValidatePutty = True Then
			AddVerificationInfoToResult  "Info" , "Invoice uploaded on WCC"
			CaptureSnapshot()
		else
			AddVerificationInfoToResult  "Error" , "Invoice could not be uploaded on WCC"
			CaptureSnapshot()
			iModuleFail = 1
			Exit Function
		End If

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If
End Function




'#################################################################################################
' 	Function Name : Disconnection_fn
' 	Description : 
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function Disconnection_fn()


			call SetObjRepository ("Putty",sProductDir & "Resources\")
		
		
		
			Dim adoData	  
			strSQL = "SELECT * FROM [BulkCSV$] WHERE  [RowNo]=" & iRowNo
			Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
			
			sType = adoData( "Type")  & ""
			sContent = adoData( "Content")  & ""
			sIndex = adoData( "Index")  & ""
			sMSISDN = DictionaryTest_G("ACCNTMSISDN")
'sMSISDN = "447389711551"
			sdate = Now
			sdate1 = Replace(sdate," ","_")
			sdate2 = Replace(sdate1,"/","_")
			sdateNew = Replace(sdate2,":","_")
			sFileName = "EDW_BULK_REQ_"&sdateNew&".csv"
		
		
			iIndex = 0
		
			If sContent<>"" Then
				If sType = "Single Disconnection"  Then
					sNewContent = Replace(sContent,"MSISDN",sMSISDN)
					AddVerificationInfoToResult  "Pass" , "New Content : "&sNewContent

						Window("PuTTY").Activate
				
						Window("PuTTY").PuttyType  "$","cd /opt/crm/siebel/batchfs/bulk/EDWDiscon/input"
			'			Window("PuTTY").PuttyType  "$","vi EDW_BULK_REQ_xyz.csv"
						Window("PuTTY").PuttyType  "$","vi "&sFileName
						Window("PuTTY").PuttyType "",micIns
						Window("PuTTY").PuttyType "",micBack
						Window("PuTTY").PuttyType  "",sNewContent
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",micEsc
						wait 1
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",":wq"
						Window("PuTTY").PuttyType  "","cd /opt/crm/siebel/batchfs/scripts"
			
						Window("PuTTY").PuttyType  "",". ./Global_Variables.sh"
						sValidatePutty = ValidatePutty ("Global_Variables.sh","1Pa55word")
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "GLobal Variables shell script executed successfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "GLobal Variables shell script not executed successfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If
			
						Window("PuTTY").PuttyType  "",". ./vf_bulk_process_EDWDiscon.sh"
						wait(5)
						sValidatePutty = ValidatePutty ("EDWDiscon.sh","FILE  '"&sFileName&"' COMPLETED")     '''FILE  'EDW_BULK_REQ_xy111z.csv' COMPLETED
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "EDW Disconnection file completed sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "EDW Disconnection file completed unsucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If	
			
								sValidatePutty = ValidatePutty ("EDWDiscon.sh"," BULK INBOUND BTACH PROCESS COMPLETED")     
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "BULK INBOUND completed sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "BULK INBOUND completed unsucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If
						End If

				Else
					 iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description

			End If

			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If
			
		End Function



'#################################################################################################
' 	Function Name :BulkCSV_fn
' 	Description : Performs all Bulk CSV functions
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function BulkCSV_fn()


			call SetObjRepository ("Putty",sProductDir & "Resources\")
		
		
		
			Dim adoData	  
			strSQL = "SELECT * FROM [BulkCSV$] WHERE  [RowNo]=" & iRowNo
			Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
			
			sType = adoData( "Type")  & ""
			sContent = adoData( "Content")  & ""
			sIndex = adoData( "Index")  & ""
			sFileName = adoData("FileName")  & ""
			sPath = adoData("Path")  &""
			sShellScript = adoData("ShellScript")  & ""
			sMSISDN = DictionaryTest_G("ACCNTMSISDN")
'sMSISDN = "447389711551"
			sdate = Now
			sdate1 = Replace(sdate," ","_")
			sdate2 = Replace(sdate1,"/","_")
			sdateNew = Replace(sdate2,":","_")
'			sFileName = "EDW_BULK_REQ_"&sdateNew&".csv"
			sFileName = sFileName&sdateNew&".csv"
			sSUP02Path = adoData("SUP02Path")  & ""
		
			iIndex = 0
		
		'	If sContent<>"" Then
			
					sNewContent = Replace(sContent,"MSISDN",sMSISDN)
					AddVerificationInfoToResult  "Pass" , "New Content : "&sNewContent

						Window("PuTTY").Activate

                         If sEnv = "E4" Then
                                      Window("PuTTY").PuttyType  "$",""&sSUP02Path
                        Else
                                      Window("PuTTY").PuttyType  "$",""&sPath
                         End If
											
						
						Window("PuTTY").PuttyType  "$","vi "&sFileName
						Window("PuTTY").PuttyType "",micIns
						Window("PuTTY").PuttyType "",micBack
						Window("PuTTY").PuttyType  "",sNewContent
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",micEsc
						wait 1
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",":wq"
						   If sEnv = "E4" Then
                                Window("PuTTY").PuttyType  "","cd /opt/SP/oracle/crm/batchfs/scripts"
                           Else
                                    Window("PuTTY").PuttyType  "","cd /opt/crm/siebel/batchfs/scripts"
                           End If

			
						Window("PuTTY").PuttyType  "",". ./Global_Variables.sh"
						sValidatePutty = ValidatePutty ("Global_Variables.sh","1Pa55word")
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "GLobal Variables shell script executed successfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "GLobal Variables shell script not executed successfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If

						CaptureSnapshot()
						Window("PuTTY").PuttyType  "",""&sShellScript
						wait(5)
						sValidatePutty = ValidatePutty (""&sShellScript,"FILE  '"&sFileName&"' COMPLETED")     '''FILE  'EDW_BULK_REQ_xy111z.csv' COMPLETED
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "File Saved Sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "File Not Saved Sucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If	
			
								sValidatePutty = ValidatePutty ("process.sh"," PROCESS COMPLETED")     
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "Process Completed Sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "Process Completed Sucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If
						

			'	Else
'					 iModuleFail = 1
'					AddVerificationInfoToResult  "Error" , Err.Description

		'	End If

			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If
			
		End Function


'#################################################################################################
' 	Function Name :CaptureLogs_fn()
' 	Description : Performs all Bulk CSV functions
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function CaptureLogs_fn()


			call SetObjRepository ("Putty",sProductDir & "Resources\")
		
		
		
			Dim adoData	  
			strSQL = "SELECT * FROM [CaptureLogs$] WHERE  [RowNo]=" & iRowNo
			Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
			
			sType = adoData( "Type")  & ""
			sContent = adoData( "Content")  & ""
			sIndex = adoData( "Index")  & ""
			sFileName = adoData("FileName")  & ""
			sPath = adoData("Path")  &""
			sShellScript = adoData("ShellScript")  & ""
			sMSISDN = DictionaryTest_G("ACCNTMSISDN")

			sFileName = sFileName&sdateNew&".csv"
		
		
			iIndex = 0
	
			
					sNewContent = Replace(sContent,"MSISDN",sMSISDN)
					AddVerificationInfoToResult  "Pass" , "New Content : "&sNewContent

						Window("PuTTY").Activate
				
'						Window("PuTTY").PuttyType  "$","cd /opt/crm/siebel/batchfs/bulk/input"
						Window("PuTTY").PuttyType  "$",""&sPath
						Window("PuTTY").PuttyType  "$","vi "&sFileName
						Window("PuTTY").PuttyType "",micIns
						Window("PuTTY").PuttyType "",micBack
						Window("PuTTY").PuttyType  "",sNewContent
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",micEsc
						wait 1
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",":wq"
						Window("PuTTY").PuttyType  "","cd /opt/crm/siebel/batchfs/scripts"
			
						Window("PuTTY").PuttyType  "",". ./Global_Variables.sh"
						sValidatePutty = ValidatePutty ("Global_Variables.sh","1Pa55word")
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "GLobal Variables shell script executed successfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "GLobal Variables shell script not executed successfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If

						CaptureSnapshot()
						Window("PuTTY").PuttyType  "",""&sShellScript
						wait(5)
						sValidatePutty = ValidatePutty (""&sShellScript,"FILE  '"&sFileName&"' COMPLETED")     '''FILE  'EDW_BULK_REQ_xy111z.csv' COMPLETED
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "File Saved Sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "File Not Saved Sucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If	
			
								sValidatePutty = ValidatePutty ("process.sh"," PROCESS COMPLETED")     
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "Process Completed Sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "Process Completed Unsucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If
						

			'	Else
'					 iModuleFail = 1
'					AddVerificationInfoToResult  "Error" , Err.Description

		'	End If

			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If
			
		End Function


'#################################################################################################
' 	Function Name :DirectDebitForPutty_fn
' 	Description : Performs all Bulk CSV functions
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function DirectDebitForPutty_fn()


			call SetObjRepository ("Putty",sProductDir & "Resources\")
		
		
		
			Dim adoData	  
'			strSQL = "SELECT * FROM [BulkCSV$] WHERE  [RowNo]=" & iRowNo
'			Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
			
			
						Window("PuTTY").Activate
				
					 If sEnv = "SUP02" Then
                                Window("PuTTY").PuttyType  "","cd /opt/SP/oracle/crm/batchfs/scripts"
                           Else
                                    Window("PuTTY").PuttyType  "","cd /opt/crm/siebel/batchfs/scripts"
                           End If
						   		
						Window("PuTTY").PuttyType  "",". ./Global_Variables.sh"
						sValidatePutty = ValidatePutty ("Global_Variables.sh","1Pa55word")
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "GLobal Variables shell script executed successfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "GLobal Variables shell script not executed successfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If

						CaptureSnapshot()
						Window("PuTTY").PuttyType  "",". ./VF_DirectDebit.sh"
						wait(5)
		
			
								sValidatePutty = ValidatePutty ("DirectDebit.sh","WORKFLOW  COMPLETED")     
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "Direct Debit Workflow Completed Sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "Direct Debit Workflow  Completed Unsucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If


			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If
			
		End Function


'#################################################################################################
' 	Function Name :ADDACSAUDDIS_fn
' 	Description : Performs all Bulk CSV functions
'   Created By :  Payel
'	Creation Date :        
'##################################################################################################
Public Function ADDACSAUDDIS_fn()


			call SetObjRepository ("Putty",sProductDir & "Resources\")
		
		
		
			Dim adoData	  
			strSQL = "SELECT * FROM [ADDACSAUDDIS$] WHERE  [RowNo]=" & iRowNo
			Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)
			
			sType = adoData( "Type")  & ""
			sContent = adoData( "Content")  & ""
			sIndex = adoData( "Index")  & ""
			sFileName = adoData("FileName")  & ""
			sPath = adoData("Path")  &""
			sShellScript = adoData("ShellScript")  & ""
			sMSISDN = DictionaryTest_G("ACCNTMSISDN")
			sMandateID =DictionaryTest_G.Item("MandateID")
'sMSISDN = "447389711551"
			sdate = Now
			sdate1 = Replace(sdate," ","_")
			sdate2 = Replace(sdate1,"/","_")
			sdateNew = Replace(sdate2,":","_")
'			sFileName = "EDW_BULK_REQ_"&sdateNew&".csv"
			sFileName = sFileName&sdateNew&".csv"
		
		
			iIndex = 0
		
		'	If sContent<>"" Then
			
					sNewContent = Replace(sContent,"MandateID",sMandateID)
					AddVerificationInfoToResult  "Info" , "New Content : "&sNewContent

						Window("PuTTY").Activate
				
'						Window("PuTTY").PuttyType  "$","cd /opt/crm/siebel/batchfs/bulk/input"
						Window("PuTTY").PuttyType  "$",""&sPath
						Window("PuTTY").PuttyType  "$","vi "&sFileName
						Window("PuTTY").PuttyType "",micIns
						Window("PuTTY").PuttyType "",micBack
						Window("PuTTY").PuttyType  "",sNewContent
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",micEsc
						wait 1
						CaptureSnapshot()
						Window("PuTTY").PuttyType "",":wq"
						Window("PuTTY").PuttyType  "","cd /opt/crm/siebel/batchfs/scripts"
			
						Window("PuTTY").PuttyType  "",". ./Global_Variables.sh"
						sValidatePutty = ValidatePutty ("Global_Variables.sh","1Pa55word")
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "GLobal Variables shell script executed successfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "GLobal Variables shell script not executed successfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If

						CaptureSnapshot()
						Window("PuTTY").PuttyType  "",""&sShellScript
						wait(5)
						sValidatePutty = ValidatePutty (""&sShellScript,"FILE  '"&sFileName&"' COMPLETED")     '''FILE  'EDW_BULK_REQ_xy111z.csv' COMPLETED
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "File Saved Sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "File Not Saved Sucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If	
			
								sValidatePutty = ValidatePutty ("process.sh"," PROCESS COMPLETED")     
								If sValidatePutty = True Then
									AddVerificationInfoToResult  "Info" , "Process Completed Sucessfully"
									CaptureSnapshot()
								else
									AddVerificationInfoToResult  "Error" , "Process Completed Sucessfully"
									CaptureSnapshot()
									iModuleFail = 1
									Exit Function
								End If
						

			'	Else
'					 iModuleFail = 1
'					AddVerificationInfoToResult  "Error" , Err.Description

		'	End If

			If Err.Number <> 0 Then
					iModuleFail = 1
					AddVerificationInfoToResult  "Error" , Err.Description
			End If
			
		End Function
