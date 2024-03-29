'#################################################################################################
' 	Function Name : WinSelect
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WinSelect(obj,data)
    If data<>"[Do Nothing]" Then
		If obj.Exist(30) Then
			obj.Select data
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" doesnot exist"
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
	End If
End Function

RegisterUserFunc "WinTreeView", "WinSelect", "WinSelect"


'#################################################################################################
' 	Function Name : WinSet
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WinSet(obj,data)
    If data<>"[Do Nothing]" Then
		If obj.Exist(30) Then
			obj.Set data
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" doesnot exist"
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
	End If
End Function

RegisterUserFunc "WinEdit", "WinSet", "WinSet"


'#################################################################################################
' 	Function Name : WinRadioSet
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WinRadioSet(obj)

		If obj.Exist(30) Then
			obj.Set
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" doesnot exist"
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

End Function
RegisterUserFunc "WinRadioButton", "WinRadioSet", "WinRadioSet"



'#################################################################################################
' 	Function Name : WinClick
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WinClick(obj)

		If obj.Exist(30) Then
			obj.Highlight
			obj.Click
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" doesnot exist"
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

End Function
RegisterUserFunc "WinButton", "WinClick", "WinClick"


'#################################################################################################
' 	Function Name : PuttyType
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function PuttyType(obj,TypeAgainst,TypeValue)
   If TypeValue<>"" Then
		TypeAgainst1 = Replace(TypeAgainst,"Exist:","")
		sResult = ValidatePuttyData (TypeAgainst1,"LastLine")
		If sResult =  True Then
			obj.Type TypeValue
			obj.Type micReturn
		elseif instr ( 1, TypeAgainst, "Exist:",1)>0 Then

			 AddVerificationInfoToResult "Info", TypeAgainst & " not found to enter " & TypeValue
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", TypeAgainst & " not found to enter " & TypeValue
			 ExitAction("FAIL")
		End If
	
	
		If Err.Number <> 0 Then
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", Err.Description
			 Call UpdateStepStatusInResult("Fail")
			 Call CaptureScreenshot
			  ExitAction("FAIL")
		End If
   End If
End Function

RegisterUserFunc "Window", "PuttyType", "PuttyType"



'#################################################################################################
' 	Function Name : PuttyTypeOnly
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function PuttyTypeOnly(obj,TypeAgainst,TypeValue)
   If TypeValue<>"" Then

		sResult = ValidatePuttyData (TypeAgainst,"LastLine")
		If sResult =  True Then
			obj.Type TypeValue
		elseif instr ( 1, TypeAgainst, "OPTIONAL",1)>0 Then

			 AddVerificationInfoToResult "Info", TypeAgainst & " not found to enter " & TypeValue
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", TypeAgainst & " not found to enter " & TypeValue
			 ExitAction("FAIL")
		End If
	
	
		If Err.Number <> 0 Then
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", Err.Description
			 Call UpdateStepStatusInResult("Fail")
			 Call CaptureScreenshot
			  ExitAction("FAIL")
		End If
   End If
End Function

RegisterUserFunc "Window", "PuttyTypeOnly", "PuttyTypeOnly"


Public Function ValidatePuttyData(data,where)
   iCnt = 0
   flag = False
'   Set oFSOPuttyLog = CreateObject("Scripting.FileSystemObject")
'	Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
	Do
		Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
		If where = "LastLine" Then
			Do Until oFilePuttyLog.AtEndOfStream
'				On error resume next
			   lineNextStr = oFilePuttyLog.ReadLine
'			   err.clear
			Loop
			If instr(1,lineNextStr,data,1)>0 Then
				flag = True
				ValidatePuttyData = True
'				oFSOPuttyLog.Close
				Exit Function
			End If
		End If
'		oFSOPuttyLog.Close
		iCnt = iCnt +1
'		print "waited for ValidatePuttyData" & 2*iCnt
'		print lineNextStr
		wait 10
	Loop While (flag = False and iCnt<20)

End Function



Function Kill_OpenPuttyInstances()
	 strSQL = "Select * From Win32_Process Where Name Like 'putty.exe%'"
		 
	Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set ProcColl = oWMIService.ExecQuery(strSQL)
				 
	For Each oElem in ProcColl
	oElem.Terminate
	Next


End Function

Public Function ValidatePuttyData1(data,where)
   iCnt = 0
   flag = False
'   Set oFSOPuttyLog = CreateObject("Scripting.FileSystemObject")
'	Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
	Do

		Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
		If where = "LastLine" Then
			Do Until oFilePuttyLog.AtEndOfStream
			   lineNextStr = oFilePuttyLog.ReadLine
			Loop
			If instr(1,lineNextStr,data,1)>0 Then				
				flag = True
				ValidatePuttyData = True
				oFilePuttyLog.Close
				Exit Function
			End If
		elseIf where = "" Then
			
			 lineNextStr = oFilePuttyLog.ReadAll
			
			If instr(1,lineNextStr,data,1)>0 Then				
				flag = True
				ValidatePuttyData = True
				Exit Function
			End If
		End If
		iCnt = iCnt +1
		wait 2
	Loop While (flag = False and iCnt<20)
	ValidatePuttyData = False
End Function


Public Function CapturePuttyData(data,key1,sIndex)
	iCnt = 0
	flag = False
'	actIndex=0
	DictionaryTest_G(key1) = ""
'   Set oFSOPuttyLog = CreateObject("Scripting.FileSystemObject")
'	Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
	Do
			actIndex=0
			Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
			Do Until oFilePuttyLog.AtEndOfStream
				 lineNextStr = oFilePuttyLog.ReadLine
				 If data <> "LastLine" Then
				
					data1 = Replace(data,"LINEAFTER:","")
					If instr(1,lineNextStr,data1,1)>0 Then
						If instr(1,data,"LINEAFTER:",1)>0 Then	
							 lineNextStr = oFilePuttyLog.ReadLine
							 If lineNextStr="" Then
								 lineNextStr = oFilePuttyLog.ReadLine
							 End If
						End If
						If Cstr(actIndex)=Cstr(sIndex) Then
							flag = True
							CapturePuttyData = True
							DictionaryTest_G(key1) = lineNextStr
							AddVerificationInfoToResult  "Info" , key1 & " is " & lineNextStr
	'						oFilePuttyLog.Close
							Exit Function
						else
							actIndex = actIndex + 1
						End If
					End If
				End If
			Loop
			If data = "LastLine" Then
				flag = True
				CapturePuttyData = True
				DictionaryTest_G(key1) = lineNextStr
				AddVerificationInfoToResult  "Info" , key1 & " is " & lineNextStr
'				oFilePuttyLog.Close
				Exit Function
			End If
		iCnt = iCnt +1
		wait 10
'		oFilePuttyLog.Close
	Loop While (flag = False and iCnt<70)
	CapturePuttyData = False
	iModuleFail = 1
	AddVerificationInfoToResult  "Error" , data & " is not found"
End Function

Public Function ValidatePutty(where,data)
   iCnt = 0
   flag = False
'   Set oFSOPuttyLog = CreateObject("Scripting.FileSystemObject")
'	Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
	Do
			Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
				If where = "" Then
				
					 lines = oFilePuttyLog.ReadAll
					 If instr (1,lines,data,1)>0 Then
						 flag = True
						ValidatePutty = True
						AddVerificationInfoToResult  "Pass" , "Expected data found " & data
						Exit Function
					 End If
				else
					lines = oFilePuttyLog.ReadAll
					 If InstrRev (lines,data,-1,1)>0 Then
						If InstrRev (lines,data,-1,1) > instr (1,lines,where,1)  Then
						
							 flag = True
							ValidatePutty = True
							AddVerificationInfoToResult  "Pass" , "Expected data found " & data & " after " & where
							Exit Function
						End If
					 End If

				End If

		iCnt = iCnt +1
		wait 10
'		oFilePuttyLog.Close
	Loop While (flag = False and iCnt<50)
	ValidatePutty = False
	iModuleFail = 1
	AddVerificationInfoToResult  "Error" , "Expected data not found " & data
End Function


Public Function TypeCapturePuttyData(obj,where,what,data,key1)
   iCnt = 0
   flag = False

	Do
'			
			obj.PuttyType where,what
			Set oFilePuttyLog = oFSOPuttyLog.OpenTextFile(sLogPath,1,True)
'			
'			print iCnt
			Do  Until oFilePuttyLog.AtEndOfStream
'				wait 1
'				On error resume next
				 lineNextStr = oFilePuttyLog.ReadLine
'				 err.clear
'				 print lineNextStr
				 If data <> "LastLine" Then
				
					data1 = Replace(data,"LINEAFTER:","")
					If instr(1,lineNextStr,data1,1)>0 Then		
						If instr(1,data,"LINEAFTER:",1)>0 Then	
							 lineNextStr = oFilePuttyLog.ReadLine
						End If
						flag = True
						TypeCapturePuttyData = True
						DictionaryTest_G(key1) = lineNextStr
						AddVerificationInfoToResult  "Info" , key1 & " is " & lineNextStr

						Exit Function
					End If
				End If
			Loop
			If data = "LastLine" Then
				flag = True
				TypeCapturePuttyData = True
				DictionaryTest_G(key1) = lineNextStr
				AddVerificationInfoToResult  "Info" , key1 & " is " & lineNextStr
				Exit Function
			End If
		iCnt = iCnt +1
		wait 10
	Loop While (flag = False and iCnt<10)
'	iModuleFail = 1
'	AddVerificationInfoToResult  "Error" , data & " is not found"
	TypeCapturePuttyData = False
End Function

Public Function CreateLogFile(name1)
   Set oFSOLog = CreateObject ("Scripting.FileSystemObject")
   Set oFSOLogFile = oFSOLog.GetFile (sLogPath)

    Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789" 
   For i=1 to 10
	   str=str & mid (LETTERS,RandomNumber(1,Len(LETTERS)),1)

   Next

	lineNo= instrRev(sLogPath,"\")
	sFolderPath = Mid(sLogPath,1,lineNo-1)
	oFSOLogFile.Copy sFolderPath & "\" & name1 & "_" & str & ".log"
'	oFSOLogFile.Delete
'	oFSOLog.CreateTextFile sLogPath,True

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Const ForWriting = 2
	Set objFile = objFSO.OpenTextFile(sLogPath, ForWriting)
	
	objFile.Write ""


End Function




