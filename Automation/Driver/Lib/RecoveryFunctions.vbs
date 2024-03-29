
Function RecoveryFunction1(Object, Method, Arguments, retVal)
	Dim sDesc,ObjectType,sNames,Args
	sDesc=DescribeResult(retVal)
	ObjectType = Object.GetROProperty("micClass")	
	'Print Err.Description
	sName = Object.GetROProperty("name")
	'Print Err.Description
	sArgs =  Arguments(0)
	'Print Err.Description
	For i=1 to Ubound(Arguments) -1
		sArgs = sArgs & ", " & Arguments(i)
	Next	 
	'Print Err.Description

	AddVerificationInfoToResult  "Error" , sDesc & " " & ObjectType & " " & sName

	iModuleFail = 1
	Call UpdateStepStatusInResult("Fail")
	Call AddFailureDescToResult( sDesc,ObjectType,sName,Method,sArgs)
	Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
	Services.EndTransaction "Start"
	
	'Added to display remaining steps==========================
	Do While sCurrent_Value <> ""
		Err.Clear
		iCol = iCol + 1
		sModule = "module" & icol
		sCurrent_Value = adoRS(sModule)
		If sCurrent_Value <> "" Then
			sMyArray = split(sCurrent_Value,";",-1,1)
			sCurrent_Module = sMyArray(0)
			If Ubound(sMyArray) <> 0 Then
				iRowNo = sMyArray(1)
			End If

			Call AddStepToResult(cstr(iCol),sCurrent_Module,"Not Executed")

			'Print "Error : " & Err.Description

			If Ubound(sMyArray) <> 0 Then
				call GetDataUsed()
			End If
			Call UpdateStepStatusInResult("Not Executed")
			'Print "Error : " & Err.Description
			'Print sCurrent_Module
		End If
	Loop

	'================================
'	ExitAction(2)
End Function
