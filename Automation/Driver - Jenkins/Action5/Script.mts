
If ExecutionJourney<>"SecondPart" Then

	 iModuleFail =0			'Resetting the test case status to not fail

	sTestCaseName = adoRS("TestCaseName") & ""
	sTCID = adoRS("TCID") & ""


	Call CreateTCFolder (sTestCaseName)
'	xmlReportTCPath= CreateTCXMLResultfile(sTestCaseName)

	Services.StartTransaction "Start"
	Services.LogMessage("============================Start of the TestCase '" & sTestCaseName &"'===============================================" )



	Call AddTestCaseToResult(sTCID,sTestCaseName    ,"Not Executed") ''Reporting to xml
'	Call AddTestCaseToTCResultFile(sTCID,sTestCaseName    ,"Not Executed") ''Reporting to xml
   
	'==========Start of - Executing the Different Modules in a Test Case =============
	
	 iCol = 1
	sModule = "module" & iCol
	 sCurrent_Value = adoRS(sModule)
	
	Do While sCurrent_Value <> "" and sCurrent_Value<>"SECONDPART"
		Err.Clear
		sMyArray = split(sCurrent_Value,";",-1,1)
		sCurrent_Module = sMyArray(0)

		If Ubound(sMyArray) <> 0 Then
			iRowNo = sMyArray(1)
		End If
		If Ubound(sMyArray) = 1 Then
			iNeg = 0
		ElseIf Ubound(sMyArray) <> 0 Then
			iNeg = sMyArray(2)	
		End If
	 
		Call AddStepToResult(cstr(iCol),sCurrent_Module,"Not Executed")
'		Call AddStepToTCResultFile(cstr(iCol),sCurrent_Module,"Not Executed")
		If Ubound(sMyArray) <> 0 Then
			call GetDataUsed()
		End If
		'Print "sDataFileName " & Environment("sDataFileName")
		'Print "sDataSheetName " & Environment("sDataSheetName")

		Services.LogMessage("Executing the " & sCurrent_Module &" function")

		'sPar contains text that gets dis[played on screen - modify to include/exclude details
		sPar = "Executing:" & sCurrent_Module &  "=|||=Test_Step:" & iCol & "=|||=Testcase:" & Replace(Trim(sTestCaseName), " ", "_") & "=|||=TCID:" & sTCID & "=|||=TestProgress:" & Environment("iTestCnt") & "/" & Cstr(adoRS.RecordCount)
		'Call Display_fn(sPar)
		
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

		End If	
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Reporting  Test Step Result'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		If iModuleFail=0 Then
			Call UpdateStepStatusInResult("Pass")
'			Call UpdateStepStatusInTCResultFile("Pass")
'			Call AddPassDescToResult
'			Call AddPassDescToTCResultFile

		ElseIf iModuleFail=2 Then
			Call UpdateStepStatusInResult("Fail")
'			Call UpdateStepStatusInTCResultFile("Fail")
			Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
			Services.EndTransaction "Start"

			
			ExitAction(1)
		Else
			Call UpdateStepStatusInResult("Fail")
'			Call UpdateStepStatusInTCResultFile("Fail")
			Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
			Services.EndTransaction "Start"
			Call AddFailureDescToResult( sDesc,ObjectType,sName,Method,sArgs)
'			Call AddFailureDescToTCResultFile( sDesc,ObjectType,sName,Method,sArgs)
			
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
else

	 iModuleFail =0			'Resetting the test case status to not fail

	sTestCaseName = adoRS("TestCaseName") & ""
	sTCID = adoRS("TCID") & ""
	Call RetrieveOrderDetailsFromOutputXcelFile_fn()
	If sTCResult<>"PASS-FirstPart" and sTCResult<>"Not Executed" and sTCResult<>"Pass" Then
		ExitAction("Fail")
	elseif sTCResult = "Not Executed" Then
		ExitAction("Not Executed")
	End If

'	Call CreateTCFolder (sTestCaseName)
'	xmlReportTCPath= CreateTCXMLResultfile(sTestCaseName)

	Services.StartTransaction "Start"
	Services.LogMessage("============================Start of the TestCase '" & sTestCaseName &"'===============================================" )



'	Call AddTestCaseToResult(sTCID,sTestCaseName    ,"Not Executed") ''Reporting to xml
'	Call AddTestCaseToTCResultFile(sTCID,sTestCaseName    ,"Not Executed") ''Reporting to xml
   
	'==========Start of - Executing the Different Modules in a Test Case =============
	iCol =1
	secondPartAvailable="N"
	sModule = "module" & iCol
	sCurrent_Value = adoRS(sModule)
	Do While sCurrent_Value <> ""
		sModule = "module" & iCol
		 sCurrent_Value = adoRS(sModule)
		 If sCurrent_Value="SECONDPART" Then
			 Call AddStepToResult(cstr(iCol),sCurrent_Value,"Pass")
			 iCol=iCol+1
			 secondPartAvailable="Y"
			 
			 Exit Do
		 End If
		 iCol=iCol+1
	Loop

	If secondPartAvailable="N" Then
		ExitAction("Pass")
	End If


	sModule = "module" & iCol
	sCurrent_Value = adoRS(sModule)

	Do While sCurrent_Value <> ""
		Err.Clear
		sMyArray = split(sCurrent_Value,";",-1,1)
		sCurrent_Module = sMyArray(0)

		If Ubound(sMyArray) <> 0 Then
			iRowNo = sMyArray(1)
		End If
		If Ubound(sMyArray) = 1 Then
			iNeg = 0
		ElseIf Ubound(sMyArray) <> 0 Then
			iNeg = sMyArray(2)	
		End If
		Call AddStepToResult(cstr(iCol),sCurrent_Module,"Not Executed")
'		Call AddStepToResultPostSubmission(cstr(iCol),sCurrent_Module,"Not Executed")
'		Call AddStepToTCResultFile(cstr(iCol),sCurrent_Module,"Not Executed")
		If Ubound(sMyArray) <> 0 Then
			call GetDataUsed()
		End If
		'Print "sDataFileName " & Environment("sDataFileName")
		'Print "sDataSheetName " & Environment("sDataSheetName")

		Services.LogMessage("Executing the " & sCurrent_Module &" function")

		'sPar contains text that gets dis[played on screen - modify to include/exclude details
		sPar = "Executing:" & sCurrent_Module &  "=|||=Test_Step:" & iCol & "=|||=Testcase:" & Replace(Trim(sTestCaseName), " ", "_") & "=|||=TCID:" & sTCID & "=|||=TestProgress:" & Environment("iTestCnt") & "/" & Cstr(adoRS.RecordCount)
		'Call Display_fn(sPar)
		
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

		End If	
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Reporting  Test Step Result'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		If iModuleFail=0 Then
			Call UpdateStepStatusInResult("Pass")
'			Call UpdateStepStatusInTCResultFile("Pass")
			Call AddPassDescToResult
'			Call AddPassDescToTCResultFile

		ElseIf iModuleFail=2 Then
			Call UpdateStepStatusInResult("Fail")
'			Call UpdateStepStatusInTCResultFile("Fail")
			Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
			Services.EndTransaction "Start"
			
			ExitAction(1)
		Else
			Call UpdateStepStatusInResult("Fail")
'			Call UpdateStepStatusInTCResultFile("Fail")
			Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
			Services.EndTransaction "Start"
			Call AddFailureDescToResult( sDesc,ObjectType,sName,Method,sArgs)
'			Call AddFailureDescToTCResultFile( sDesc,ObjectType,sName,Method,sArgs)
			
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






End If