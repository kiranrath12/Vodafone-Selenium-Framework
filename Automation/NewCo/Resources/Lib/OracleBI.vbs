'#################################################################################################
' 	Function Name : LoginToOracleBI_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function LoginToOracleBI_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [LoginToOracleBI$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\OracleBI.xls",strSQL)

	sURL = adoData("URL"&sEnv) & ""
	sUsername = adoData("Username") & ""
	sPassword = adoData("Password") & ""
	
	'sOR
	call SetObjRepository ("OracleBI",sProductDir & "Resources\")
    Call CloseAllBrowsers
	systemUtil.Run "IEXPLORE.EXE",sURL,"","Open",3
	Browser("Oracle BI").Page("Oracle BI").WebEdit("id").WebSet sUsername
	Browser("Oracle BI").Page("Oracle BI").WebEdit("passwd").WebSet sPassword
	Browser("Oracle BI").Page("Oracle BI").WebButton("Sign In").WebClick

	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : OpenBRMInvoicing_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function OpenBRMInvoicing_fn()

'Get Data
'----
	'sOR
	Dim adoData	  
    strSQL = "SELECT * FROM [OpenBRMInvoicing$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\OracleBI.xls",strSQL)

	sType = adoData("Type") & ""

	call SetObjRepository ("OracleBI",sProductDir & "Resources\")

	Browser("Oracle BI").Page("Oracle BI").Link("Home").WebClick
	poid=DictionaryTest_G("PoID")
'	poid="1785686539"
	If poid = "" Then
		iModuleFail = 1
		AddVerificationInfoToResult "Error" , "PoID is empty"
		Exit Function
	End If
	Browser("Oracle BI").Page("Oracle BI").Link("Catalog Folders").WebClick


	If Browser("Oracle BI").Page("Oracle BI").WebTable("Shared Folders").Image("disclosure_collapsed").Exist(2) Then
		Browser("Oracle BI").Page("Oracle BI").WebTable("Shared Folders").Image("disclosure_collapsed").WebClick
	End If

	If Browser("Oracle BI").Page("Oracle BI").WebTable("BRM_Invoices").Image("disclosure_collapsed").Exist(2) Then
		Browser("Oracle BI").Page("Oracle BI").WebTable("BRM_Invoices").Image("disclosure_collapsed").WebClick
	End If
	
	Browser("Oracle BI").Page("Oracle BI").WebElement("0.0.0.1").WebClick


'	poid="1630151094"
	resFolder=Environment("RunFolderPath") & "\" & Environment("ResFldrName")
	Browser("Oracle BI").Page("Oracle BI").Link("InvoiceOpen").WebClick

	Browser("Oracle BI").Page("Oracle BI").WebEdit("_paramsp_Inv_MIN_POID").WebSet poid
	Browser("Oracle BI").Page("Oracle BI").WebEdit("_paramsp_Inv_MAX_POID").WebSet poid
	
	Browser("Oracle BI").Page("Oracle BI").WebButton("Apply").WebClick
	Browser("index:=0").Page("index:=0").Sync
	Browser("Oracle BI").Page("Oracle BI").Link("html tag:=A","innerhtml:="&sType).WebClick
'	cnt = 0
'	Do While (Browser("Oracle BI").Page("Oracle BI").Image("Loading ...").Exist(0) and cnt<10)
'		Wait 1
'		cnt =cnt +1 
'	Loop
	wait 3
	text = Browser("Oracle BI").Page("Oracle BI").ActiveX("Adobe Acrobat DC Browser").WinObject("AVPageView").GetVisibleText

	If instr(text,DictionaryTest_G("BillNo")) > 0 Then
		AddVerificationInfoToResult  "Info" , "PDF generated with invoice no " & DictionaryTest_G("BillNo")
	else
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If
	CaptureSnapshot()

	Browser("Oracle BI").Page("Oracle BI").Image("Actions").WebClick
	Browser("Oracle BI").Page("Oracle BI").Link("Export").WebClick
	Browser("Oracle BI").Page("Oracle BI").Link("PDF").WebClick
'	Browser("Oracle BI").Dialog("File Download").WinButton("Save").Highlight
	Browser("Oracle BI").Dialog("File Download").WinButton("Save").WinClick
'	print TCFolderPath & "\" & poid
	Dialog("Save As").WinToolbar("ToolbarWindow32").Highlight
	Dialog("Save As").WinToolbar("ToolbarWindow32").Press 1
	Dialog("Save As").WinEdit("Edit_2").WinSet resFolder
	Dialog("Save As").WinEdit("Edit_2").Type  micReturn 

	wait 2
	fileName = Dialog("Save As").WinEdit("Edit").GetRoProperty ("text")

'	Dialog("Save As").WinEdit("Edit").Highlight
'	Dialog("Save As").WinEdit("Edit").Click
'	Dim mySendKeys
'	set mySendKeys = CreateObject("WScript.shell")
'	mySendKeys.SendKeys poid
'	Dialog("Save As").WinEdit("Edit").Type  micReturn
	
'	Dialog("Save As").WinComboBox("ComboBox").Highlight
'	Dialog("Save As").WinComboBox("ComboBox").Select "All Files"
'	Dialog("Save As").WinComboBox("ComboBox").Select "All Files"
'	Dialog("Save As").WinComboBox("ComboBox").WinComboSelect "All Files"
'	Dialog("Save As").WinEdit("Edit").WinSet "C:\BRM_Bursting_Invoice_Report_PostPaid_invoice1.pdf"

	wait 2
	Dialog("Save As").WinButton("Save").WinClick
	If	Dialog("Download complete").WinButton("Close").Exist(0) Then
		Dialog("Download complete").WinButton("Close").WinClick
	End If
	wait 2
	Set FSO = CreateObject("Scripting.FileSystemObject")
	strFile = resFolder & "\" & fileName & ".pdf"
	strRename = resFolder & "\" & poid & ".pdf"

	If FSO.FileExists(strFile) Then
		FSO.MoveFile strFile, strRename
	End If

	Set FSO = Nothing
'	Dialog("Download complete").WinButton("Close").WinClick
'
	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function
