	 Call Display_fn("QTP_is_starting_a_test-case_run.....")
	 iModuleFail =0			'Resetting the test case status to not fail

	sTestCaseName = adoRS("TestCaseName") & ""
	sTCID = adoRS("TCID") & ""

	Services.StartTransaction "Start"
	Services.LogMessage("============================Start of the TestCase '" & sTestCaseName &"'===============================================" )


	'======================Code For Dependent Cases=======================
'	arrDependentOn=Split(adoRS("DependentOn") & "",chr(10))
'	For i = 0 to Ubound(arrDependentOn)
'		 If  Dictionary_G.Exists(arrDependentOn(i)) Then
'			 Call AddTestCaseToResult(sTCID,sTestCaseName    ,"Not Executed-Dependent On: "  & adoRS("DependentOn") & "") ''Reporting to xml
'			 Dictionary_G.Add  sTestCaseName ,"Not Executed"
'			 ExitAction("Not Executed")
'		End if
'	Next
	'======================================================================

	Call AddTestCaseToResult(sTCID,sTestCaseName    ,"Not Executed") ''Reporting to xml
   
	'==========Start of - Executing the Different Modules in a Test Case =============
	
	 iCol = 36
	sModule = "module" & iCol
	 sCurrent_Value = adoRS(sModule)
	
	Do While sCurrent_Value <> ""
		Err.Clear
		sMyArray = split(sCurrent_Value,";",-1,1)
		sCurrent_Module = sMyArray(0)
		iRowNo = sMyArray(1)
		If Ubound(sMyArray) = 1 Then
			iNeg = 0
		Else
			iNeg = sMyArray(2)	
		End If
	 
		Call AddStepToResult(cstr(iCol),sCurrent_Module,"Not Executed")
		If InStr(sCurrent_Module,"_co")>0 Then 'sCurrent_Module is a function
			 sCurrent_Module = "ClickOn_fn"
		End if
		If InStr(sCurrent_Module,"_es")>0 Then 'sCurrent_Module is a function
			 sCurrent_Module = "ExecuteStatements_fn"
		End if
		If InStr(sCurrent_Module,"_voe")>0 Then 'sCurrent_Module is a function
			 sCurrent_Module = "VerifyObjectExist_fn"
		End if
		call GetDataUsed()
		'Print "sDataFileName " & Environment("sDataFileName")
		'Print "sDataSheetName " & Environment("sDataSheetName")

		Services.LogMessage("Executing the " & sCurrent_Module &" function")

		'sPar contains text that gets dis[played on screen - modify to include/exclude details
		sPar = "Executing:" & sCurrent_Module &  "=|||=Test_Step:" & iCol & "=|||=Testcase:" & Replace(Trim(sTestCaseName), " ", "_") & "=|||=TCID:" & sTCID & "=|||=TestProgress:" & Environment("iTestCnt") & "/" & Cstr(adoRS.RecordCount)
		Call Display_fn(sPar)
		
		If InStr(sCurrent_Module,"[")=0 Then 'sCurrent_Module is a function
			iModuleFail =0
			If InStr(sCurrent_Module,"(")=0 Then    
				Execute ( "Call " & sCurrent_Module & "()" )
			Else
				Execute ( "Call " & sCurrent_Module)
			End If	
		Else  'sCurrent_Module is a action
			RunAction sCurrent_Module, oneIteration, iRowNo, iNeg, iModuleFail
		End If

		If Err.Number <> 0 Then ' if the module was NOT found, due to any reason - not added as an external action OR spelling mistake in the Data Table
			 iModuleFail = 1
			 systemutil.CloseProcessByName("tooltip.exe")
		End If	
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Reporting  Test Step Result'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		If iModuleFail=0 Then
			Call UpdateStepStatusInResult("Pass")
			systemutil.CloseProcessByName("tooltip.exe")
		ElseIf iModuleFail=2 Then
			Call UpdateStepStatusInResult("Fail")
			Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
			Services.EndTransaction "Start"
			systemutil.CloseProcessByName("tooltip.exe")
			ExitAction(1)
		Else
			Call UpdateStepStatusInResult("Fail")
			Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
			Services.EndTransaction "Start"
			Call AddFailureDescToResult( sDesc,ObjectType,sName,Method,sArgs)
			systemutil.CloseProcessByName("tooltip.exe")
			ExitAction(1)
		End if 
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		iCol = iCol + 1
		sModule = "module" & icol
		sCurrent_Value = adoRS(sModule)
		
	Loop
	
	Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
	Services.EndTransaction "Start"
	ExitAction("Pass")

	'****************************************************************************************
	'Function: Display_fn()
	'Summary: Display testcase execution progress
	'Author: Anunay
	'Dated: 15/5/2012
	'****************************************************************************************
	Function Display_fn(par)
		systemutil.CloseProcessByName("tooltip.exe")
		systemutil.Run "C:\QA_Auto_QTP\Driver\tooltip.exe", par
	End Function














































































































































































