
''*********************************************************************************************
'Action Name: TestSetExec   AddFailureDescToResult AddVerificationInfoToResult UpdateStepStatusInResult AddFailureDescToResult ExecuteDBQuery_fn OrderFormEntry_fn econfig_fn OrdersPayment_fn

'Author: Automation Team CaptureDataOutputSheet_fn RetrieveMSISDN_fn RetrieveAccount_fn Econfig_fn CaptureOrderDetails_fn CaptureDataOutputSheet_fn OrderFormEntry_fn;2
'CaptureSnapshot  ExecExcQry UpdateContactPhoneNumber_fn
'Start Date: Feb 2015 UsedProductServices_fn  CaptureDataOutputSheet_fn ProductServicesVerify_fn;3
'AddProduct_fn
'Brief description: This Action is used to execute the list of Test Cases MenuLineItems_fn;1 ExecuteDBQuery_fn;2 ExecExcQry VerifyAccountSummary_fn MenuLineItems_fn AddProduct_fn RetrieveMSISDN_fn RetrieveMSISDN_fn SelectExpandRow DirectDebit_fn LoginToSiebel_fn CustomiseProduct_fn AddProduct_fn


Call ClearErrorInVbs_fn()
Call  InitializeVariables()	
Call setpath(Environment.Value("Product")) 


Dim xmlReportPath, iTestCnt
Dim sPrefixXmlResultFileWith

sPrefixXmlResultFileWith = Environment.Value("ResultFileName")

'For Reporting results to xml
'*********************************************************************************************
xmlReportPath= CreateAXMLResultfile(sPrefixXmlResultFileWith)

'Query for executing test cases*********************************************************************************************
sTestCaseList = Environment.Value("TestCaseFile") 


strSQL =Environment.Value("Query")
Set adoRS = ExecExcQry(sTestCaseList , strSQL)

If  adoRS.Eof  Then ' No test case seleted for execution
	ExitAction
End If

Call UpdateResultSummaryTotal(Cstr(adoRS.RecordCount))

iTestCnt = 0		'Count of executed testcases in a module
adoRS.MoveFirst
'ExeuctionJourney="FirstPart"
'
Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("User")
objEnv("TCResult")=""
Do While Not adoRS.EOF
	For rerun = 1 To 2
		
		
		Call  InitializeDictionaryTest()
	'    Set objShell = CreateObject("WScript.Shell")
	'    Set objEnv = objShell.Environment("User")
	 
	    outputRowNum=objEnv("outputRowNum")+1
		objEnv("outputRowNum")=outputRowNum
	'    Environment("outputRowNum")=Environment("outputRowNum")+1
		iTestCnt = iTestCnt + 1
		Environment("iTestCnt") = iTestCnt
		sTCResult = 0
		sTestCaseName=""
		Call ClearErrorInVbs_fn()
		Call SetUpDataProxy()
	   
		sTCResult = RunAction ("ExecuteTestCases", oneIteration)
		
		Set vDataProxyConnector=Nothing
	
		If sTCResult = "Pass"   Then
			sTCResult = "PASS-FirstPart"
			Call UpdateTestCaseStatusInResult("PASS-FirstPart")
			Reporter.ReportEvent micPass, sTestCaseName, "All modules of the" & sTestCaseName & " passed "
			iContinuousFailNumber = 0
		Else
			If  sTCResult<>"Not Executed" Then
				Call UpdateTestCaseStatusInResult("FAIL") 
				IsPreTCFailed ="Yes"
				Dictionary_G.Add  sTestCaseName ,"FAIL"
				iContinuousFailNumber = iContinuousFailNumber +1
			End If
		End if
	
		Call WriteToOutputXcelFile_fn()
	   
		iNumberOfTestCasesExecuted = iNumberOfTestCasesExecuted +  1
		If sTCResult <> "PASS-FirstPart" Then
			iAtLeastOneTestFailed =  iAtLeastOneTestFailed + 1
			If iContinuousFailNumber > 5 Then
	'			SendAlerts()
				iContinuousFailNumber = 0
	'		Exit Do
			End If
		else
			Exit For
		End if
	Next
	adoRS.MoveNext
Loop
'
'Set adoRS = ExecExcQry(sTestCaseList , strSQL)
'
'If  adoRS.Eof  Then ' No test case seleted for execution
'	ExitAction
'End If
'
'Call UpdateResultSummaryTotal(Cstr(adoRS.RecordCount))
'
'iTestCnt = 0		'Count of executed testcases in a module
ExecutionJourney="SecondPart"
adoRS.MoveFirst

Do While Not adoRS.EOF
	Call  InitializeDictionaryTest()
	iTestCnt = iTestCnt + 1
	Environment("iTestCnt") = iTestCnt
	sTCResult = 0
	sTestCaseName=""
	Call ClearErrorInVbs_fn()
	Call SetUpDataProxy()
   
	sTCResult = RunAction ("ExecuteTestCases", oneIteration)
	If sTCResult <> "Pass"   Then
		sTCResult = "Fail"
	End If
	sAccountNo=""
	If DictionaryTest_G.Exists("AccountNo") Then
		sAccountNo=DictionaryTest_G("AccountNo")
	End If
	sOrderNo=""
	If DictionaryTest_G.Exists("OrderNo") Then
		sOrderNo=DictionaryTest_G.Item("OrderNo")
	End If
	objEnv("TCResult")=objEnv("TCResult") & sTestCaseName & ":" & sTCResult & ":" & sAccountNo & ":" & sOrderNo & ";"
	Set vDataProxyConnector=Nothing

	If sTCResult = "Pass"   Then
		Call UpdateTestCaseStatusInResult("PASS")
		Reporter.ReportEvent micPass, sTestCaseName, "All modules of the" & sTestCaseName & " passed "
		iContinuousFailNumber = 0
	Else
		If  sTCResult<>"Not Executed" Then
			Call UpdateTestCaseStatusInResult("FAIL") 
			IsPreTCFailed ="Yes"
			Dictionary_G.Add  sTestCaseName ,"FAIL"
			iContinuousFailNumber = iContinuousFailNumber +1
		End If
	End if

	Call UpdateOutputXcelFile_fn()
   
	iNumberOfTestCasesExecuted = iNumberOfTestCasesExecuted +  1
'	If sTCResult <> "Pass" Then
		iAtLeastOneTestFailed =  iAtLeastOneTestFailed + 1
'		If iContinuousFailNumber > 5 Then
''			SendAlerts()
'			iContinuousFailNumber = 0
''		Exit Do
'		End If
'	End if
	adoRS.MoveNext
Loop


adoRS.Close
Set adoRS = Nothing

If  iAtLeastOneTestFailed <> 0  Then
	Reporter.ReportEvent 1, "Total failures", "The number of test cases that failed were ' " & iAtLeastOneTestFailed  & " '"   
End If
Call WriteResultSummary()

