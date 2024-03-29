'#################################################################################################
' 	Function Name : WebClick
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WebClick(obj)
	If obj.Exist(120) Then
		obj.Click
		AddVerificationInfoToResult "Pass", obj.ToString &" is clicked"
	else
		iModuleFail = 1
		 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
		 Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 ExitAction("FAIL")
		'err.raiseFAIL
	End If

End Function

RegisterUserFunc "Image", "WebClick", "WebClick"
RegisterUserFunc "Link", "WebClick", "WebClick"

'#################################################################################################
' 	Function Name : WebSet
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WebSet(obj,data)
	If obj.Exist(120) Then
		obj.Set data
		AddVerificationInfoToResult "Pass", obj.ToString &" is set with " & data
	else
		 iModuleFail = 1
		 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
		Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 Call RemainingTestCaseSteps
		  ExitAction("FAIL")
	End If

End Function

RegisterUserFunc "WebEdit", "WebSet", "WebSet"
RegisterUserFunc "WebCheckBox", "WebSet", "WebSet"


'#################################################################################################
' 	Function Name : WebSet
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WebSelect(obj,data)

	If obj.Exist(120) Then
		obj.Select data
		AddVerificationInfoToResult "Pass", obj.ToString &" is set with " & data
	else
		 iModuleFail = 1
		 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
		Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 Call RemainingTestCaseSteps
		  ExitAction("FAIL")
	End If

End Function

RegisterUserFunc "WebList", "WebSelect", "WebSelect"



'#################################################################################################
' 	Function Name : CloseAllBrowsers
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CloseAllBrowsers
	SystemUtil.CloseProcessByName "iexplore.exe"

End Function


'#################################################################################################
' 	Function Name : WebGetRoProperty
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WebGetRoProperty(obj,data)


		If obj.Exist(180) Then
			WebGetRoProperty=obj.GetRoProperty (data)
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" does not exist"
			 Call UpdateStepStatusInResult("Fail")
		     Call CaptureScreenshot
		      ExitAction("FAIL")
		End If
		If Err.Number <> 0 Then
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", Err.Description
			 Call UpdateStepStatusInResult("Fail")
		     Call CaptureScreenshot
		      ExitAction("FAIL")
		End If
		WebGetRoProperty = False

	
End Function

RegisterUserFunc "WebElement", "WebGetRoProperty", "WebGetRoProperty"





