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
	End Function

	Public Function SetUpDataProxy()
		Set vDataProxyConnector = CreateObject("Newco.DataServices.Proxy")
				
		If (vDataProxyConnector.ContextSetup(sEnv)) Then
'			Print "Context initialized: " & sEnv
		Else
			ScriptFaulted"Error Context NOT initialized"
		End If
	End Function




















































Extern.declare micUInteger, "RegisterWindowMessage", "user32.dll", "RegisterWindowMessageA", micString
Extern.declare micLResult, "SendMessage", "user32.dll", "SendMessageA", micHwnd, micUInteger, micWParam, micLParam
Extern.Declare micLong, "GetWindowThreadProcessId", "user32.dll", "GetWindowThreadProcessId", micHwnd, micLong
Extern.Declare micLong, "GetCurrentThreadId", "kernel32.dll", "GetCurrentThreadId"
Extern.Declare micShort, "AttachThreadInput", "user32.dll", "AttachThreadInput", micDWord, micDWord, micShort
Extern.Declare micLong, "GetCursor", "user32.dll", "GetCursor"
Extern.Declare micLong, "LoadCursorA", "user32.dll", "LoadCursor", micLong, micLong
Extern.Declare micLong, "GetForegroundWindow", "user32.dll", "GetForegroundWindow"
Extern.Declare micInteger, "PostMessage",  "user32.dll", "PostMessageW", micHwnd, micUInteger, micWParam, micLParam
Extern.Declare micHwnd, "GetActiveWindow", "user32.dll", "GetActiveWindow"


    UNICA_SLB_COPYTEXT = extern.RegisterWindowMessage("UNICA_SLB_COPYTEXT")
	UNICA_SLB_SETCURSEL = extern.RegisterWindowMessage("UNICA_SLB_SETCURSEL")
	UNICA_SLB_EXPANDALL = extern.RegisterWindowMessage("UNICA_SLB_EXPANDALL")
	UNICA_SLB_COLLAPSEALL = extern.RegisterWindowMessage("UNICA_SLB_COLLAPSEALL")
	UNICA_SLB_GETTREESEP = extern.RegisterWindowMessage("UNICA_SLB_GETTREESEP")
	UNICA_SLB_GETITEMCOUNT = extern.RegisterWindowMessage("UNICA_SLB_GETITEMCOUNT")

	UNICA_SLB_FINDSTRING = extern.RegisterWindowMessage("UNICA_SLB_FINDSTRING")
	UNICA_SLB_FINDSTRINGEXACT = extern.RegisterWindowMessage("UNICA_SLB_FINDSTRINGEXACT")

	UNICA_PB_GETFIRSTGREYEDNAME = extern.RegisterWindowMessage("UNICA_PB_GETFIRSTGREYEDNAME")
	UNICA_PB_SELECTBYNAME = extern.RegisterWindowMessage("UNICA_PB_SELECTBYNAME")
	UNICA_PB_POPUPCONTEXTMENU = extern.RegisterWindowMessage("UNICA_PB_POPUPCONTEXTMENU")
	UNICA_PB_CONNECT = extern.RegisterWindowMessage("UNICA_PB_CONNECT")
