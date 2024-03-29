




  Public Function SetObjRepository (sObjRep,sObjRepDir) 
	
	Err.Clear
	Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
	Dim qtRepositories 'As QuickTest.Resources ' Declare a Resources object variable

	Set qtApp =CreateObject("QuickTest.Application") 
    Set qtRepositories = qtApp.Test.Actions("ExecuteTestCases").ObjectRepositories 
	qtRepositories.RemoveAll 
	
	sObjRepFullPath = sObjRepDir & sObjRep & ".tsr"
	qtRepositories.Add sObjRepFullPath,1
	
	Set qtRepositories = Nothing ' Release the Application object
	
	Set qtApp = Nothing ' Release the Application object
	
	SetObjRepository = 1

End Function
'Ganesh
Public function SetPath(sProductName)
		Dim i
		i= InStrRev (Environment.Value("TestDir") ,"\Driver")
		sProjectDir = left(Environment.Value("TestDir"),i)		
		sProductDir = sProjectDir & sProductName & "\"
  End Function
  
  Public Function GetExcConnStr(ExcWBFilePath)
	GetExcConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&ExcWBFilePath&";Extended Properties=Excel 8.0;"

	
End Function 



Public Function ExecExcQry(ExcelWBFilPth, ExcelSQL)
	Const adUseClient = 3
	Const adOpenKeyset = 1
	Const adLockOptimistic = 3
	Const adCmdText = 1
	Const adOpenForwardOnly = 0
	Const adLockReadOnly = 1
	Const adOpenDynamic = 2
	Dim adoExcConn 'As ADODB.Connection
	Dim adoExcRS 'As ADODB.Recordset

	Dim strExcConn 'As String
	 
	strExcConn = GetExcConnStr(ExcelWBFilPth)
	
	Set adoExcConn = CreateObject("ADODB.Connection")
	Set adoExcRS = CreateObject("ADODB.Recordset")
	
	With adoExcConn
		.ConnectionString = strExcConn
		.CursorLocation = adUseClient
		.Open
	End With
'	adoExcConn.Open adoExcConn
'	adoExcConn.Execute ExcelSQL
	adoExcRS.Open ExcelSQL, adoExcConn, adOpenKeyset , adLockOptimistic
	''Newly Added Code
	If  adoExcRS.Eof  OR Err.Number <> 0 Then
		Services.LogMessage("' " & ExcelWBFilPth & " is the file'")
		Services.LogMessage(" There was an error acessing the Record in the Excel file, possibly due to the record NOT being present in the file")
		Reporter.ReportEvent micFail ,"RecordNOTFound", "RecordNOTFound" 
		If  Err.Number <> 0 Then
			Services.LogMessage ("Error Description was '" & Err.Description & "'")
		Else
			Services.LogMessage("End of File - seems to be the cause of the problem")	
		End If
		iModuleFail = 1
		
		AddVerificationInfoToResult  "Error" , Err.Description
		Call UpdateStepStatusInResult("Fail")
'		Call UpdateStepStatusInTCResultFile("Fail")
		Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
		Services.EndTransaction "Start"
'		ExitAction(1)
	End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''End of newly added code   
	Set ExecExcQry = adoExcRS
End Function


Public Function ExecExcQry1(ExcelWBFilPth, ExcelSQL)
	Const adUseClient = 3
	Const adOpenKeyset = 1
	Const adLockOptimistic = 3
	Const adCmdText = 1
	Const adOpenForwardOnly = 0
	Const adLockReadOnly = 1
	Const adOpenDynamic = 2
	Dim adoExcConn 'As ADODB.Connection
	Dim adoExcRS 'As ADODB.Recordset

	Dim strExcConn 'As String
	 
	strExcConn = GetExcConnStr(ExcelWBFilPth)
	
	Set adoExcConn = CreateObject("ADODB.Connection")
	Set adoExcRS = CreateObject("ADODB.Recordset")
	
	With adoExcConn
		.ConnectionString = strExcConn
		.CursorLocation = adUseClient
		.Open
	End With
'	adoExcConn.Open adoExcConn
'	adoExcConn.Execute ExcelSQL
	adoExcRS.Open ExcelSQL, adoExcConn, adOpenKeyset , adLockOptimistic
	''Newly Added Code
	If  adoExcRS.Eof  OR Err.Number <> 0 Then
		Services.LogMessage("' " & ExcelWBFilPth & " is the file'")
		Services.LogMessage(" There was an error acessing the Record in the Excel file, possibly due to the record NOT being present in the file")
		Reporter.ReportEvent micFail ,"RecordNOTFound", "RecordNOTFound" 
		If  Err.Number <> 0 Then
			Services.LogMessage ("Error Description was '" & Err.Description & "'")
		Else
			Services.LogMessage("End of File - seems to be the cause of the problem")	
		End If
		iModuleFail = 1
		
'		AddVerificationInfoToResult  "Error" , Err.Description
'		Call UpdateStepStatusInResult("Fail")
'		Call UpdateStepStatusInTCResultFile("Fail")
'		Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
'		Services.EndTransaction "Start"
'		ExitAction(1)
	End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''End of newly added code   
	Set ExecExcQry1 = adoExcRS
End Function


Public Function UpdateExcQry(ExcelWBFilPth, ExcelSQL)

	Dim adoExcConn 'As ADODB.Connection


	Dim strExcConn 'As String
	 
	strExcConn = GetExcConnStr(ExcelWBFilPth)
	
	Set adoExcConn = CreateObject("ADODB.Connection")

	
	adoExcConn.Open strExcConn
	adoExcConn.Execute ExcelSQL

	''Newly Added Code
	If  Err.Number <> 0 Then


			Services.LogMessage ("Error Description was '" & Err.Description & "'")

		iModuleFail = 1
		  AddVerificationInfoToResult "Error", Err.Description
		Call UpdateStepStatusInResult("Fail")
		Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
		Services.EndTransaction "Start"
'		ExitAction(1)
	End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''End of newly added code   

End Function

Function SetVal (obj, x) 

       dim y 

       y = obj.GetROProperty("value") 

       Reporter.ReportEvent micDone, "previous value", y 

       MyFuncWithParam=obj.Set (x) 

End Function 

Public Function Local_ExecExcQry(ExcelWBFilPth, ExcelSQL)
	Const adUseClient = 3
	Const adOpenKeyset = 1
	Const adLockOptimistic = 3
	Const adCmdText = 1
	Const adOpenForwardOnly = 0
	Const adLockReadOnly = 1
	Const adOpenDynamic = 2
	Dim adoExcConn 'As ADODB.Connection
	Dim adoExcRS 'As ADODB.Recordset

	Dim strExcConn 'As String
	 
	strExcConn = GetExcConnStr(ExcelWBFilPth)
	
	Set adoExcConn = CreateObject("ADODB.Connection")
	Set adoExcRS = CreateObject("ADODB.Recordset")
	
	With adoExcConn
		.ConnectionString = strExcConn
		.CursorLocation = adUseClient
		.Open
	End With
	
	adoExcRS.Open ExcelSQL, adoExcConn, adOpenKeyset , adLockOptimistic
	''Newly Added Code
	If   Err.Number <> 0 Then
		Services.LogMessage("' " & ExcelWBFilPth & " is the file'")
		Services.LogMessage(" There was an error acessing the Record in the Excel file, possibly due to the record NOT being present in the file")
		Reporter.ReportEvent micFail ,"RecordNOTFound", "RecordNOTFound" 
		If  Err.Number <> 0 Then
			Services.LogMessage ("Error Description was '" & Err.Description & "'")
		Else
			Services.LogMessage("End of File - seems to be the cause of the problem")	
		End If
		iModuleFail = 1
		Call UpdateStepStatusInResult("Fail")
		Services.LogMessage("============================End of the TestCase '" & sTestCaseName &"'===============================================" )
		Services.EndTransaction "Start"
'		ExitAction(1)
	End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''End of newly added code   



	Set Local_ExecExcQry = adoExcRS
End Function
'#################################################################################################
'	Function Name : ReplaceInData
'#################################################################################################
Public function ReplaceInData(sData)

	   sData = Replace(sData,"<&1>","aug")
		ReplaceInData = Replace(sData,"<&2>","_958")	
		   
			If Err.Number <> 0 Then
				iModuleFail = 1
				AddVerificationInfoToResult "Error", Err.Description
			End If
End Function

'#################################################################################################
'            Function Name : DeleteBrowsingHistory_fn
'            Description :  This function is used Delete Browsing History
'            Created By:  Ram
'            Creation Date:   03/07/2009
'##################################################################################################
Function DeleteBrowsingHistory_fn()
                        Dim WshShell, oExec
                        Set WshShell = CreateObject("WScript.Shell")
                        Set oExec = WshShell.Exec("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255")
						wait 20
End Function
