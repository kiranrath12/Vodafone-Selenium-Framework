
Dim IsPreTCFailed

IsPreTCFailed="Yes"
	Dim iCol
	Dim outputRowNum

	
	Dim sTCID
	Dim sTestCaseName
    Dim sUrl
	Dim sCurrent_Value
    Dim sCurrent_Module
	Dim iRowNo
	Dim iNeg
	Dim iModuleFail 
	Dim iTestFail
	Dim sTCResult
	Dim sExplorerPath
	Dim sTestRun
	Dim sRequiredType
	Dim sTestCaseList
	Dim sResultFile
	Dim sProjectDir
	Dim sProductDir
'	Dim sManagerUrl_G
	'Dim sPlanUrl_G
'	Dim sEMessageUrl_G
'	Dim sCampaignUrl_G
'	Dim sProduct_G
	Dim sDatabaseConnectionString_G
	Dim ExecutionJourney
'	Dim sModelUrl_G

	
	Dim sLoadConfigServletURL_G
	Dim sConnStrExe_G
	Dim sConnStrSys_G
	Dim sConnStrTrack_G
	Dim sMailingInstID_G
	Dim sPhysicalOLTName_G
	Dim sLinksHref_G
	Dim ExecutionDetails_G
	Set Dictionary_G = CreateObject("Scripting.Dictionary")
	Set DictionaryTest_G = CreateObject("Scripting.Dictionary")
	Set DictionaryTempTest_G = CreateObject("Scripting.Dictionary")
	DictionaryTest_G.Item("AccountNo")="abk"
	DictionaryTest_G.Item("OrderNo")="giuksi"
	DictionaryTest_G.Item("Status")="od"
	Dim sDatabaseTablePrefix_G
	Dim sConnStrPlan_G

	Dim vDataProxyConnector
	

	
	Dim adoRS '
	Dim strSQL '

	Dim oFSO
	Dim oResultdoc
	Dim oRoot
	Dim oTCResultdoc
	Dim oTCRoot
	Dim sResultFileName

	Dim TCFolderPath


	Dim iAnalysisCount
	Dim Login
	
	Dim iAtLeastOneTestFailed
	Dim iNumberOfTestCasesExecuted

	Dim Parent


	
	iAtLeastOneTestFailed = 0 ' for setting the BVT final result
	iNumberOfTestCasesExecuted = 0 


  sExplorerPath = "Iexplore"
  sProduct_G = Environment.Value("Product")
   Public Function InitializeVariables()
		arrVariables = split(Environment.Value("VariableSet"),"|")
		For i = 0 to UBound(arrVariables) 
			arrVariables1=split(arrVariables(i),"^")
			strVarValue =   chr(34) & arrVariables1(1) & chr(34)
			Execute (arrVariables1(0) & " = "  & strVarValue )
		Next
	End Function


	Public Function InitializeDictionaryTest()
		DictionaryTest_G.RemoveAll
		DictionaryTempTest_G.RemoveAll
	End Function



	Public Function SetUpDataProxy()
		Set vDataProxyConnector = CreateObject("Newco.DataServices.Proxy")
				
		If (vDataProxyConnector.ContextSetup(sEnv)) Then
'			Print "Context initialized: " & sEnv
		Else
			ScriptFaulted"Error Context NOT initialized"
		End If
	End Function



































