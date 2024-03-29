
'#################################################################################################
' 	Function Name : SiebClick
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebClick1(obj)
	If obj.Exist(180) Then
		On error resume next
		Set vDictionary=CreateObject("Scripting.Dictionary")
'		CheckForPopup(vDictionary)
'		If SiebApplication("classname:=SiebApplication").Exist(5) Then
'			SiebApplication("classname:=SiebApplication").SetTimeOut(180)
'		End If

		
		obj.Click

		
'		If SiebApplication("classname:=SiebApplication").Exist(5) Then
'			SiebApplication("classname:=SiebApplication").SetTimeOut(60)
'		End If
		Err.Clear
		
		AddVerificationInfoToResult "Pass", obj.ToString &" is clicked"
	else
		iModuleFail = 1
		 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
		 Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 ExitAction("FAIL")
		'err.raiseFAIL
	End If
	If Err.Number <> 0 Then
		 iModuleFail = 1
		 AddVerificationInfoToResult "Error", Err.Description
		 Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 ExitAction("FAIL")
	End If
End Function

'RegisterUserFunc "SiebButton", "SiebClick", "SiebClick"

'#################################################################################################
' 	Function Name : SiebClick
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebClick(obj,sPopUp)

	If obj.Exist(180) Then
		On error resume next
		Set vDictionary=CreateObject("Scripting.Dictionary")

		If sPopUp=False  Then
			SiebApplication("classname:=SiebApplication").SetTimeOut(180)
		else
			SiebApplication("classname:=SiebApplication").SetTimeOut(10)
		End If
		
		obj.Click

		
'		If SiebApplication("classname:=SiebApplication").Exist(5) Then
'			SiebApplication("classname:=SiebApplication").SetTimeOut(60)
'		End If
		vResult=CheckForPopup(vDictionary,sPopUp)
		If sPopUp <> False and not(instr(sPopUp,"OPTIONAL:")>0) Then
			expPopUp = True
			If vResult <> expPopUp Then
				AddVerificationInfoToResult "Fail", "Expected pop up - " & sPopUp & "did not occur"
				iModuleFail = 1
				ExitAction("FAIL")				
			End If
		End If
		
		SiebApplication("classname:=SiebApplication").SetTimeOut(180)
'		If SiebApplication("classname:=SiebApplication").Exist(5) Then
'			SiebApplication("classname:=SiebApplication").SetTimeOut(180)
'		End If
		Err.Clear
		
		AddVerificationInfoToResult "Pass", obj.ToString &" is clicked"
	else
		iModuleFail = 1
		 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
		 Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 ExitAction("FAIL")
		'err.raiseFAIL
	End If
	If Err.Number <> 0 Then
		 iModuleFail = 1
		 AddVerificationInfoToResult "Error", Err.Description
		 Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 ExitAction("FAIL")
	End If
End Function

RegisterUserFunc "SiebButton", "SiebClick", "SiebClick"

''***********************************************************************************************************************


Public Function SiebClick2(obj,sPopUp)

	If obj.Exist(180) Then
		On error resume next
		Set vDictionary=CreateObject("Scripting.Dictionary")

		If sPopUp=False  Then
			SiebApplication("classname:=SiebApplication").SetTimeOut(20)
		else
			SiebApplication("classname:=SiebApplication").SetTimeOut(10)
		End If
		
		obj.Click

		
'		If SiebApplication("classname:=SiebApplication").Exist(5) Then
'			SiebApplication("classname:=SiebApplication").SetTimeOut(60)
'		End If
		vResult=CheckForPopup(vDictionary,sPopUp)
		If sPopUp <> False and not(instr(sPopUp,"OPTIONAL:")>0) Then
			expPopUp = True
			If vResult <> expPopUp Then
				AddVerificationInfoToResult "Fail", "Expected pop up - " & sPopUp & "did not occur"
				iModuleFail = 1
				ExitAction("FAIL")				
			End If
		End If
		
		SiebApplication("classname:=SiebApplication").SetTimeOut(180)
'		If SiebApplication("classname:=SiebApplication").Exist(5) Then
'			SiebApplication("classname:=SiebApplication").SetTimeOut(180)
'		End If
		Err.Clear
		
		AddVerificationInfoToResult "Pass", obj.ToString &" is clicked"
	else
		iModuleFail = 1
		 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
		 Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 ExitAction("FAIL")
		'err.raiseFAIL
	End If
	If Err.Number <> 0 Then
		 iModuleFail = 1
		 AddVerificationInfoToResult "Error", Err.Description
		 Call UpdateStepStatusInResult("Fail")
		 Call CaptureScreenshot
		 ExitAction("FAIL")
	End If
End Function

RegisterUserFunc "SiebButton", "SiebClick2", "SiebClick2"



'#################################################################################################
' 	Function Name : SiebSelect
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebSelect(obj,data)
    If data<>"[Do Nothing]" Then
		If obj.Exist(600) Then
			SiebApplication("classname:=SiebApplication").SetTimeOut(180)
			Recovery.Enabled=False
			On error resume next
			
			obj.ActiveItem
			obj.Select data
			Err.Clear
			Recovery.Enabled=True
			SiebApplication("classname:=SiebApplication").SetTimeOut(180)
			AddVerificationInfoToResult "Pass", obj.ToString &" has selected " & data
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
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

RegisterUserFunc "SiebPicklist", "SiebSelect", "SiebSelect"


'#################################################################################################
' 	Function Name : SiebSelectPopup
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebSelectPopup(obj,data,sPopUp)
    If data<>"[Do Nothing]" Then
		If obj.Exist(180) Then
		On error resume next
		Set vDictionary=CreateObject("Scripting.Dictionary")

		If sPopUp=False  Then
			SiebApplication("classname:=SiebApplication").SetTimeOut(30)
		else
			SiebApplication("classname:=SiebApplication").SetTimeOut(10)
		End If
		Recovery.Enabled=False

			
			obj.ActiveItem
			obj.Select data
            vResult=CheckForPopup(vDictionary,sPopUp)
		If sPopUp <> False and not(instr(sPopUp,"OPTIONAL:")>0) Then
			expPopUp = True
			If vResult <> expPopUp Then
				AddVerificationInfoToResult "Fail", "Expected pop up - " & sPopUp & "did not occur"
				iModuleFail = 1
				ExitAction("FAIL")				
			End If
		End If
		
		SiebApplication("classname:=SiebApplication").SetTimeOut(180)
			Err.Clear
			Recovery.Enabled=True
			
			AddVerificationInfoToResult "Pass", obj.ToString &" has selected " & data
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
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

RegisterUserFunc "SiebPicklist", "SiebSelectPopup", "SiebSelectPopup"


'#################################################################################################
' 	Function Name : SiebSetText
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebSetText(obj,data)
   If data<>"[Do Nothing]" Then

		If obj.Exist(180) Then
			obj.SetText data
			AddVerificationInfoToResult "Pass", obj.ToString &" is set with " & data
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
			Call UpdateStepStatusInResult("Fail")
		     Call CaptureScreenshot
			 Call RemainingTestCaseSteps
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

RegisterUserFunc "SiebCalendar", "SiebSetText", "SiebSetText"
RegisterUserFunc "SiebText", "SiebSetText", "SiebSetText"
RegisterUserFunc "SiebTextArea", "SiebSetText", "SiebSetText"


'#################################################################################################
' 	Function Name : SiebSetCheckbox1
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebSetCheckbox1(obj,data,sPopUp)
   If data<>"[Do Nothing]" Then

		If obj.Exist(180) Then
            On error resume next
			Set vDictionary=CreateObject("Scripting.Dictionary")
			If  Ucase(data)=Ucase("Off")Then
				obj.SetOff
				AddVerificationInfoToResult "Pass", obj.ToString &" is unchecked "
			ElseIf Ucase(data)=Ucase("On")Then
				obj.SetOn
				AddVerificationInfoToResult "Pass", obj.ToString &" is checked "
			else
				 iModuleFail = 1
				 AddVerificationInfoToResult "Incorrect data is passed to " & obj.ToString
				 Call UpdateStepStatusInResult("Fail")
		     Call CaptureScreenshot
		      ExitAction("FAIL") 
			End If
            vResult=CheckForPopup(vDictionary,sPopUp)
			If sPopUp <> False and not(instr(sPopUp,"OPTIONAL:")>0) Then
				expPopUp = True
				If vResult <> expPopUp Then
					AddVerificationInfoToResult "Fail", "Expected pop up - " & sPopUp & "did not occur"
					iModuleFail = 1
					ExitAction("FAIL")				
				End If
			End If

		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
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

RegisterUserFunc "SiebCheckbox1", "SiebSetCheckbox1", "SiebSetCheckbox1"



'#################################################################################################
' 	Function Name : SiebSetCheckbox
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebSetCheckbox(obj,data)
   If data<>"[Do Nothing]" Then

		If obj.Exist(180) Then
			If  Ucase(data)=Ucase("Off")Then
				obj.SetOff
				AddVerificationInfoToResult "Pass", obj.ToString &" is unchecked "
			ElseIf Ucase(data)=Ucase("On")Then
				obj.SetOn
				AddVerificationInfoToResult "Pass", obj.ToString &" is checked "
			else
				 iModuleFail = 1
				 AddVerificationInfoToResult "Incorrect data is passed to " & obj.ToString
				 Call UpdateStepStatusInResult("Fail")
		     Call CaptureScreenshot
		      ExitAction("FAIL") 
			End If

		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
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

RegisterUserFunc "SiebCheckbox", "SiebSetCheckbox", "SiebSetCheckbox"


'#################################################################################################
' 	Function Name : SiebGetCellText
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebGetCellText(obj,colName,rowNo)
   If data<>"[Do Nothing]" Then

		If obj.Exist(180) Then
			repName=obj.GetColumnRepositoryName (colName)
			value1=obj.GetCellText (repName,rowNo)
			
			SiebGetCellText=value1
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
	End If
	
End Function

RegisterUserFunc "SiebList", "SiebGetCellText", "SiebGetCellText"


'#################################################################################################
' 	Function Name : SiebDrillDownColumn
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebDrillDownColumn(obj,colName,rowNo)
'   If data<>"[Do Nothing]" Then

		If obj.Exist(180) Then
			obj.ActivateRow rowNo
			repName=obj.GetColumnRepositoryName (colName)
			CaptureSnapshot()
			obj.DrillDownColumn repName,rowNo

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
'	End If
	
End Function

RegisterUserFunc "SiebList", "SiebDrillDownColumn", "SiebDrillDownColumn"

'#################################################################################################
' 	Function Name : SiebListRowCount
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebListRowCount(obj)


		If obj.Exist(240) Then
			SiebListRowCount=obj.rowscount
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

	
End Function

RegisterUserFunc "SiebList", "SiebListRowCount", "SiebListRowCount"


'#################################################################################################
' 	Function Name : SiebGetRoProperty
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebGetRoProperty(obj,data)


		If obj.Exist(180) Then
			SiebGetRoProperty=obj.GetRoProperty (data)
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

	
End Function

RegisterUserFunc "SiebCurrency", "SiebGetRoProperty", "SiebGetRoProperty"
RegisterUserFunc "SiebText", "SiebGetRoProperty", "SiebGetRoProperty"
RegisterUserFunc "SiebPicklist", "SiebGetRoProperty", "SiebGetRoProperty"

'#################################################################################################
' 	Function Name : SiebListProductRowNo
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebListProductRowNo(obj,item1,data1)


		If obj.Exist(180) Then
				rwNum = obj.SiebListRowCount ' Taking Row count 
	
			For loopVar = rwNum-1 to 0 Step -1 
			
				strProduct =  Trim(obj.SiebGetCellText("Product",loopVar))
			
					If strProduct = Trim(sProduct) Then ' comparing product name
			

						Exit For
					End If
			
			Next
			SiebListRowCount=obj.rowscount
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

	
End Function

RegisterUserFunc "SiebList", "SiebListProductRowNo", "SiebListProductRowNo"


'#################################################################################################
' 	Function Name : SelectProductRow
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SelectProductRow(obj,sProduct,iIndex)

		flag="N"
		If obj.Exist(180) Then
'			Call ShowMoreButton_fn()
			ieIndex=0
			rwNum = obj.RowsCount ' Taking Row count 
			obj.ActivateRow 0

			If instr(sProduct,"|")<=0 Then

				For loopVar = 0 to  rwNum-1
				
					strProduct =  Trim(obj.GetCellText("Product",loopVar))
	
						If strProduct = Trim(sProduct) Then ' comparing product name
						'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
							If Cstr(iIndex)=Cstr(ieIndex) Then
								obj.ActivateRow loopVar
								flag="Y"
								Exit For
                            else
								ieIndex=ieIndex+1
							End If
						End If	
				
				Next
			else
					sProductArrs=Split(sProduct,"|")
					For each sProductArr in sProductArrs
						For loopVar = 0 to  rwNum-1
						
							strProduct =  Trim(obj.GetCellText("Product",loopVar))
			
								If strProduct = Trim(sProductArr) Then ' comparing product name
								'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
					
									obj.ActivateRow loopVar,"Ctrl"
									flag="Y"
									Exit For
								End If	
					
						Next
					Next
			End If

			If flag="N" Then
				  iModuleFail = 1
				 AddVerificationInfoToResult "Error", sProduct &" cannot be found " & obj.toString
				 Call UpdateStepStatusInResult("Fail")
				 Call CaptureScreenshot
				  ExitAction("FAIL")
			End If
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", sProduct &" doesnot exist "
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

RegisterUserFunc "SiebList", "SelectProductRow", "SelectProductRow"

'#################################################################################################
' 	Function Name : SelectNameRow
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SelectNameRow(obj,sProduct)

		flag="N"
		If obj.Exist(180) Then
		'	Call ShowMoreButton_fn()
			rwNum = obj.RowsCount ' Taking Row count 
			obj.ActivateRow 0

			If instr(sProduct,"|")<=0 Then

				For loopVar = 0 to  rwNum-1
				
					strProduct =  Trim(obj.GetCellText("Product Name",loopVar))
	
						If instr(strProduct,sProduct) > 0 Then  ' comparing product name
						'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
			
							obj.ActivateRow loopVar
							flag="Y"
							Exit For
						End If	
				
				Next
			else
					Dim cnt
					cnt=0
					sProductArrs=Split(sProduct,"|")
					For each sProductArr in sProductArrs
						For loopVar = 0 to  rwNum-1
						
							strProduct =  Trim(obj.GetCellText("Product Name",loopVar))
			
								If instr(strProduct,sProductArr) > 0 Then ' comparing product name
								'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
									If cnt=0 Then
										obj.ActivateRow loopVar
										flag="Y"
										cnt=cnt+1
										Exit For
									else
										obj.SelectRow loopVar,"Ctrl"
										flag="Y"
										Exit For
									End If
								End If	
					
						Next

						If  flag="N" Then
							iModuleFail = 1
							AddVerificationInfoToResult "Error", sProductArr &" cannot be found " & obj.toString
							Call UpdateStepStatusInResult("Fail")
							Call CaptureScreenshot
							ExitAction("FAIL")
						End If
					Next
					
			End If

			If flag="N" Then
				  iModuleFail = 1
				 AddVerificationInfoToResult "Error", sProduct &" cannot be found " & obj.toString
				 Call UpdateStepStatusInResult("Fail")
				 Call CaptureScreenshot
				  ExitAction("FAIL")
			End If
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", sProduct &" doesnot exist "
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

RegisterUserFunc "SiebList", "SelectNameRow", "SelectNameRow"


'#################################################################################################
' 	Function Name : SelectInstalledIdRow
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SelectInstalledIdRow(obj,sProduct)

		flag="N"
		If obj.Exist(180) Then
		'	Call ShowMoreButton_fn()
			rwNum = obj.RowsCount ' Taking Row count 
			repName=obj.GetColumnRepositoryName ("Installed ID")
			obj.ActivateRow 0

			If instr(sProduct,"|")<=0 Then

				For loopVar = 0 to  rwNum-1
				
					strProduct =  Trim(obj.GetCellText(repName,loopVar))
	
						If strProduct = Trim(sProduct) Then ' comparing product name
						'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
			
							obj.ActivateRow loopVar
							flag="Y"
							Exit For
						End If	
				
				Next
			else
					sProductArrs=Split(sProduct,"|")
					For each sProductArr in sProductArrs
						For loopVar = 0 to  rwNum-1
						
							strProduct =  Trim(obj.GetCellText("Product Name",loopVar))
			
								If strProduct = Trim(sProductArr) Then ' comparing product name
								'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
					
									obj.SelectRow loopVar,"Ctrl"
									flag="Y"
									Exit For
								End If	
					
						Next
					Next
			End If

			If flag="N" Then
				  iModuleFail = 1
				 AddVerificationInfoToResult "Error", sProduct &" cannot be found " & obj.toString
				 Call UpdateStepStatusInResult("Fail")
				 Call CaptureScreenshot
				  ExitAction("FAIL")
			End If
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", sProduct &" doesnot exist "
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

RegisterUserFunc "SiebList", "SelectInstalledIdRow", "SelectInstalledIdRow"





'#################################################################################################
' 	Function Name : SelectExpandRow
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SelectExpandRow(obj,sProduct,iIndex)


		If obj.Exist(180) Then
			repName=obj.GetColumnRepositoryName ("Product")
			ieIndex=0

			rwNum = obj.RowsCount ' Taking Row count 

				For loopVar = 0 to  rwNum-1
				
					strProduct =  Trim(obj.GetCellText(repName,loopVar))
	
						If strProduct = Trim(sProduct) Then ' comparing product name
						'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
			
							If Cstr(iIndex)=Cstr(ieIndex) Then
								If not(obj.IsRowExpanded(loopVar)) Then
									obj.ActivateRow loopVar
									obj.ClickHier
								End If
								flag="Y"
								Exit For
							else
								ieIndex=ieIndex+1
							End If
						End If	
				
				Next

		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.toString &" doesnot exist "
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

RegisterUserFunc "SiebList", "SelectExpandRow", "SelectExpandRow"



'#################################################################################################
' 	Function Name : ExpandRow
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function ExpandRow(objApplet)

	Set obj = objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList")
	If obj.Exist(180) Then
		varActiveRow = obj.ActiveRow
		If not(obj.IsRowExpanded(varActiveRow)) Then			
			obj.ClickHier
		End If


	else
		 iModuleFail = 1
		 AddVerificationInfoToResult "Error", obj.toString &" doesnot exist "
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

RegisterUserFunc "SiebApplet", "ExpandRow", "ExpandRow"


'#################################################################################################
' 	Function Name : SelectCollapseRow
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SelectCollapseRow(obj,sProduct,iIndex)


		If obj.Exist(180) Then
			ieIndex=0

			rwNum = obj.RowsCount ' Taking Row count 

				For loopVar = 0 to  rwNum-1
				
					strProduct =  Trim(obj.GetCellText("Product",loopVar))
	
						If strProduct = Trim(sProduct) Then ' comparing product name
						'	SiebApplication("Siebel Call Center").SiebScreen("Orders").SiebView("Line Items View").SiebApplet("Line Items").SiebList("List").SiebDrillDownColumn sColNameInput,rwNum
			
							If Cstr(iIndex)=Cstr(ieIndex) Then
								If obj.IsRowExpanded(loopVar) Then
									obj.ActivateRow loopVar
									obj.ClickHier
								End If
								flag="Y"
								Exit For
							else
								ieIndex=ieIndex+1
							End If
						End If	
				
				Next

		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.toString &" doesnot exist "
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
RegisterUserFunc "SiebList", "SelectCollapseRow", "SelectCollapseRow"


'#################################################################################################
' 	Function Name : SiebSelectMenu
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function SiebSelectMenu(obj,data)
		flag="N"
		If obj.Exist(300) Then
			cnt=obj.count
			For i=0 to cnt-1
				repName=obj.GetRepositoryNameByIndex (i)
				If instr(1,obj.GetUIName(repName),data,1)>=1 Then
					Recovery.Enabled=False
					On error resume next
					SiebApplication("classname:=SiebApplication").SetTimeOut(180)
					obj.Select repName
					Err.clear
					Recovery.Enabled=True
					AddVerificationInfoToResult "Pass", obj.ToString &" has selected " & data
					flag="Y"
					Exit For
				End If
			Next
			If flag="N" Then
				  iModuleFail = 1
				 AddVerificationInfoToResult "Error", data &" cannot be found in SiebMenu"
				 Call UpdateStepStatusInResult("Fail")
				 Call CaptureScreenshot
				  ExitAction("FAIL")
			End If
			
		else
			 iModuleFail = 1
			 AddVerificationInfoToResult "Error", obj.ToString &" is not enabled"
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

RegisterUserFunc "SiebMenu", "SiebSelectMenu", "SiebSelectMenu"

Public Function CheckForPopup ( ByVal pDictionary,ByVal sPopUp)

	Dim vPopupCount
    Dim vMessage

    vPopupCount = 0
	If  instr(sPopUp,"OPTIONAL:") > 0 Then
		sPopUp = Replace(sPopUp,"OPTIONAL:","")
	End If

	'Check for Message From WebPage - is there a Browser of the type "windowcontext" = "appliaction
'	If  instr(sPopUp,"FRAUD RISK") > 0 Then
'		Set vBrowserPopup = Browser("windowcontext:=Application")
'	Else
'		Set vBrowserPopup = Browser("windowcontext:=Application").Dialog("micclass:=Dialog","text:=Message from webpage")
'	End If
	Set vBrowserPopup = Browser("windowcontext:=Application").Dialog("micclass:=Dialog","text:=Message from webpage")
	
    If ( vBrowserPopup.Exist(0))Then
		vPopupCount = vPopupCount + 1
		REM Call GetLogger().LogEvent ("Message From Webpage PopUp:-")
		pDictionary("Outcome") = "Message From WebPage"
		REM Call GetLogger().StoreSnapshotA( vBrowserPopup, "Message from webpage")
		If vBrowserPopup.Static( "window id:=65535").Exist(0)Then
			vMessage = vBrowserPopup.Static("window id:=65535").GetROProperty("Text")
			pDictionary("Message") = vMessage
			Call ValidatePopUpMsg (sPopUp,vMessage)
'			If instr(sPopUp,"|") Then
'				sPopUpArr = Split(sPopUp,"|")
'			End If
'			If instr(vMessage,sPopUp) > 0  Then
'					AddVerificationInfoToResult "Info","Expected popup-." & vMessage
'					CaptureSnapshot()
'			else
'				AddVerificationInfoToResult "Fail","Unexpected popup-." & vMessage
'				 iModuleFail = 1
'				 CaptureSnapshot()
'			End If
			
			Call HandleWebPagePopup( vMessage, vBrowserPopup )
			REM Call GetLogger().LogEvent ("Message From Webpage:" & vMessage)
			REM Call SetOutputParams ( "Popup", "Err: Message From Webpage" )
		End If
		REM Call HandleWebPagePopup( vMessage, vBrowserPopup )
	End if

	'check fo SiebPopups
	 If  ( CheckForSiebelPopups ( "Application", pDictionary,sPopUp )) Then
		vPopupCount = vPopupCount + 1
	 End if
	If  ( CheckForSiebelPopups ( "Popup", pDictionary,sPopUp )) Then
		vPopupCount = vPopupCount + 1
	 End if
	If (CheckForSiebelMessage ( "", pDictionary,sPopUp)) Then
	   vPopupCount = vPopupCount + 1
	End If

	If ( vPopupCount > 0 )Then
		
		CheckForPopup  pDictionary,sPopUp
		CheckForPopup  = True
	Else
		CheckForPopup = False
	End If
	    
End Function		

Public Function ValidatePopUpMsg (sPopUp,vMessage)
			Dim sPopUpArr
			flag = "N"
'			If instr(sPopUp,"|") Then
				sPopUpArr = Split(sPopUp,"|")
'			End If
			For i=0 to Ubound(sPopUpArr)
				If instr(vMessage,sPopUpArr(i)) > 0  or instr(vMessage,"Please review the following messages before proceeding")>0Then
					AddVerificationInfoToResult "Info","Expected popup-." & vMessage
					CaptureSnapshot()
					flag = "Y"
					Exit For
				End If
			Next
			
			If flag = "N" Then
				AddVerificationInfoToResult "Fail","Unexpected popup-." & vMessage
				 iModuleFail = 1
				 CaptureSnapshot()
			End If

End Function
		
		
Sub HandleWebPagePopup ( pMessage, vPopup )
   'To do : change behaviour based on mesage for know popups, currently if OK exists will be clicked otherwise pop up will be closed
   'To do: Winobject library
   'To do; Globally log the occutance of known popup types e.g Siebel Server Busy
		If  vPopup.WinButton("text:=OK").Exist(0)Then
			vPopup.WinButton("text:=OK").Click
		Else
			vPopup.Close
		End If
End Sub

Public Function CheckForSiebelPopups ( pType, ByVal pDictionary,sPopUp )
   '[1]  - 2 known types - a) "windowcontext = Application, b "windowcontext = Application
    Dim vPopUp
	Dim vResult

	vResult = False
	Set vPopUp = Browser("windowcontext:=" & pType ).Dialog("micclass:=Dialog","text:=Siebel")
    If( vPopUp.Exist(0))Then
		 vResult = True
		'Log that there is a Sieb pop up
		 REM Call GetLogger().LogEvent ("Siebel PopUp:-")
		 REM Call GetLogger().StoreSnapshotA( vPopup, "Siebel Popup")
		   'is a sieb pop up so st ore outcome to sieb popup
			pDictionary("Outcome") = "Siebel Popup"
		  'if static message extists
			If( vPopUp.Static("window id:=65535").Exist(0))Then
				'retrieve the message
				vMessage = vPopUp.Static("window id:=65535").GetROProperty("Text")
				'store in dictionary
				pDictionary("Message") = vMessage
				Call ValidatePopUpMsg (sPopUp,vMessage)
'                If instr(vMessage,sPopUp)>0 Then
'					AddVerificationInfoToResult "Info","Expected popup-." & vMessage
'					CaptureSnapshot()
'				else
'					AddVerificationInfoToResult "Fail","Unexpected popup-." & vMessage
'					 iModuleFail = 1
'					 CaptureSnapshot()
'				End If
				'log the message
				REM Call GetLogger().LogEvent ("Siebel PopUp:" & vMessage)
				
				REM Call SetOutputParams ( "Popup", "Err: Siebel PopUp"  )
				Call HandleSiebelPopup ( vMessage, vPopUp )
			End if
			REM Call HandleSiebelPopup ( vMessage, vPopUp )
		End if
		CheckForSiebelPopups = vResult
End Function        		

Sub HandleSiebelPopup ( pMessage, pPopUp )
	'To do change behaviour for known popup types
	'Globally log known error types
	'currently just close the popup
	pPopUp.Close
End Sub



Public Function CheckForSiebelMessage ( pApplet, ByVal pDictionary,sPopUp )
   '[1]  - 2 known types - a) "windowcontext = Application, b "windowcontext = Application
    Dim vPopUp
	Dim vResult

	vResult = False
	Set vPopUp = SiebApplication("micclass:=SiebApplication" ).SiebScreen("micclass:=SiebScreen").SiebView("micclass:=SiebView").SiebApplet("micclass:=SiebApplet", "repositoryname:=UMF Messages Active Popup Applet")
    If( vPopUp.Exist(0))Then
		 vResult = True
		'Log that there is a Sieb pop up
		 REM Call GetLogger().LogEvent ("Siebel Messages:-")
		 REM Call GetLogger().StoreSnapshotA( vPopup, "Siebel Messages")
		   'is a sieb pop up so st ore outcome to sieb popup
			pDictionary("Outcome") = "Siebel Messages"
		  'if static message extists
			If( vPopUp.SiebTextArea( "micclass:=SiebTextArea", "repositoryname:=Message Text").Exist(0))Then
				'retrieve the message
				vMessage = vPopUp.SiebTextArea( "micclass:=SiebTextArea", "repositoryname:=Message Text").GetROProperty("Text")
				'store in dictionary
				pDictionary("Message") = vMessage
				Call ValidatePopUpMsg (sPopUp,vMessage)
'                 If instr(vMessage,sPopUp)>0 or instr(vMessage,"Please review the following messages before proceeding")>0 Then
'					AddVerificationInfoToResult "Info","Expected popup-." & vMessage
'					CaptureSnapshot()
'				else
'					AddVerificationInfoToResult "Fail","Unexpected popup-." & vMessage
'					 iModuleFail = 1
'					 CaptureSnapshot()
'				End If
				Call HandleSiebelMessages ( vMessage, vPopUp )
				'log the message
				REM Call GetLogger().LogEvent ("Siebel PopUp:" & vMessage)
				REM Call SetOutputParams ( "Popup", "Err: Siebel PopUp"  )
			End if
'			Call HandleSiebelMessages ( vMessage, vPopUp )
		End if
		CheckForSiebelMessage = vResult
End Function        		

Sub HandleSiebelMessages ( pMessage, pPopUp )
	'To do change behaviour for known popup types
	'Globally log known error types
	'currently just close the popup

	'[1] KNOWN TYPE - COMMITMENT HAS BEEN BROKEN - CLICK ACCEPT
'	If ( Instr(pMessage, "A commitment has been broken") <> 0) Then
'		pPopUp.SiebButton("micclass:=SiebButton", "repositoryname:=Response1").Click
'	Else
		pPopUp.SiebButton("micclass:=SiebButton", "repositoryname:=CloseApplet").Click
'	End If

	 
	
	
End Sub

Public Function UpdateSiebList (objApplet,objDetails,value1)
   Set obj = objApplet.SiebList ("micClass:=SiebList", "repositoryname:=SiebList")
	className = Split(objDetails,"|")(0)
	uiName = Split(objDetails,"|")(1)
	flag="N"
'	If oDict.Exists (uiName) Then
		repName=obj.GetColumnRepositoryName (uiName)
		If className = "List"  Then
			If obj.SiebPicklist ("micClass:=SiebPicklist", "repositoryname:="&repName).Exist Then
				obj.SiebPicklist ("micClass:=SiebPicklist", "repositoryname:="&repName).SiebSelect value1
				flag="Y"
			 End If
		elseif className = "Text"  Then
			If obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
				 obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).SiebSetText value1
				 flag="Y"
			End If
		elseif className = "TextArea"  Then
			If obj.SiebTextArea ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
				 obj.SiebTextArea ("micClass:=SiebText", "repositoryname:="&repName).SiebSetText value1
				 flag="Y"
			End If
		elseif className = "OpenPopUp"  Then
	
			If obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
                SiebApplication("classname:=SiebApplication").SetTimeOut(10)

				Recovery.Enabled=False
				 obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).OpenPopup

                 SiebApplication("classname:=SiebApplication").SetTimeOut(180)
				Err.Clear
				Recovery.Enabled=True
				 flag="Y"
			End If
		elseif className = "CheckBox"  Then
			If obj.SiebText ("micClass:=SiebCheckbox", "repositoryname:="&repName).Exist Then
				 obj.SiebText ("micClass:=SiebCheckbox", "repositoryname:="&repName).SiebSetCheckbox value1
				 flag="Y"
			End If
		elseif className = "CaptureTextValue"  Then
			If obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
				 varActiveRow = obj.ActiveRow
				 value1 = obj.SiebGetCellText (uiName, varActiveRow)
				 flag="Y"
			End If
		elseif className = "CaptureTextValue1"  Then
'			If obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
				 varActiveRow = obj.ActiveRow
				 value1 = obj.SiebGetCellText (uiName, varActiveRow)
				 flag="Y"
'			End If
		elseif className = "CaptureCompareTextValue"  Then
			If obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
				If DictionaryTest_G.Exists(value1) Then
					value1 = Replace (value1,value1,DictionaryTest_G(value1))
				End If
				 varActiveRow = obj.ActiveRow
				 actValue1 = obj.SiebGetCellText (uiName, varActiveRow)
				 If instr(value1,"*") Then
					value1 = Replace (value1,"*","")
					If instr(1,actValue1, value1, 1) > 0 Then
						 flag="Y"
						 AddVerificationInfoToResult "Pass", "Expected value " & value1 & "* match with actual " & actValue1
					else
						flag="N"
						AddVerificationInfoToResult "Fail", "Expected value " & value1 & "* doesnot match with actual " & actValue1
					End If
				else
					If Ucase(actValue1) = Ucase(value1) Then
						 flag="Y"
						 AddVerificationInfoToResult "Pass", "Expected value " & value1 & "* match with actual " & actValue1
					else
						flag="N"
						AddVerificationInfoToResult "Fail", "Expected value " & value1 & "* doesnot match with actual " & actValue1
					End If
				 End If
				 flag="Y"
				 '''------------------------------------Added by  Pushkar-------------------------------------------------
			End If
					elseif className = "CaptureCompareTextNegative"  Then
			If obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
				If DictionaryTest_G.Exists(value1) Then
					value1 = Replace (value1,value1,DictionaryTest_G(value1))
				End If
				 varActiveRow = obj.ActiveRow
				 actValue1 = obj.SiebGetCellText (uiName, varActiveRow)
				 If instr(value1,"*") Then
					value1 = Replace (value1,"*","")
					If instr(1,actValue1, value1, 1) > 0 Then
						 flag="N"
						 AddVerificationInfoToResult "Fail", "Expected value " & value1 & "* match with actual " & actValue1
					else
						flag="Y"
						AddVerificationInfoToResult "Pass", "Expected value " & value1 & "* doesnot match with actual " & actValue1
					End If
				else
					If Ucase(actValue1) = Ucase(value1) Then
						 flag="N"
						 AddVerificationInfoToResult "Fail", "Expected value " & value1 & "* match with actual " & actValue1
					else
						flag="Y"
						AddVerificationInfoToResult "Pass", "Expected value " & value1 & "* doesnot match with actual " & actValue1
					End If
				 End If
				 flag="Y"
			End If
			''' ----------------------------------------End Keyword --------------------------------------------
		elseif className = "DrillDown"  Then
			If obj.SiebText ("micClass:=SiebText", "repositoryname:="&repName).Exist Then
				 varActiveRow = obj.ActiveRow
				 CaptureSnapshot()
				 obj.SiebDrillDownColumn uiName, varActiveRow
				 flag="Y"
			End If
		End If
	If  flag="N" Then
		iModuleFail = 1
		AddVerificationInfoToResult "Error",uiName & "column not found"
	End If
'	else
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",uiName & "column not found"
'	End If

End Function

Public Function LocateColumns1(obj,oDict)

	records=obj.RecordCounter
	cnt = obj.ColumnsCount
	If cnt = 0  Then
		AddVerificationInfoToResult "Error","No columns found"
		RecordColumns = False
		Exit Function
	End If

	For i=0 to cnt-1
		repName= obj.GetColumnRepositoryNameByIndex(i)
		uiName = obj.GetColumnUIName(repName)
		oDict.Add uiName, i
	
'		print uiName
	Next
	RecordColumns = True
End Function



Public Function LocateColumns(obj,sLocateCol,sLocateColValue,index)
	parsedRows=0
	currParsedRows=0
	If instr(sLocateCol,"^") >0  Then
		LocateColumns = LocateAndSelectMultiColumns(obj,sLocateCol,sLocateColValue,index)
	elseIf instr(sLocateColValue,"All|") >0  Then
		LocateColumns = LocateColumnsAll(obj,sLocateCol,sLocateColValue,index)
	else
	
		Set oDict = CreateObject("Scripting.Dictionary")

		If instr(sLocateColValue,"Promotion")>0 Then
			sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
		End If
		arr1 = Split (sLocateCol,"|")
		arr2 = Split (sLocateColValue,"|")
		arrCnt1 = Ubound (arr1)
		For k=0 to arrCnt1
			oDict.Add arr1(k),arr2(k)	
		Next
	
		Set oDictTemp = CreateObject("Scripting.Dictionary")
	
		Set objSiebList = obj.SiebList("micclass:=SiebList","repositoryname:=SiebList")
	
		cnt = objSiebList.ColumnsCount
		If cnt = 0  Then
			AddVerificationInfoToResult "Error","No columns found"
			LocateColumns = False
			Exit Function
		End If
	
'		For i=0 to cnt-1
'			repName= objSiebList.GetColumnRepositoryNameByIndex(i)
'			uiName = objSiebList.GetColumnUIName(repName)
'			oDictTemp.Add uiName, i
'		
'	'		print uiName
'		Next
		arr =  oDict.Keys
		arrCount= Ubound(arr)
		For j=0 to arrCount
'			If oDictTemp.Exists (arr(j)) = False Then
'				AddVerificationInfoToResult "Error",key1 & "column not found"
'				LocateColumns = False
'				Exit Function
'	
'			End If
			oDictTemp (arr(j)) = j
		Next
	
		currRecord = obj.RecordCounter
		actIndex = 0
		Do
	
			getRowCount currRecord,firstRow,lastRow,totalRow
			If totalRow = 0 and sLocateColValue="Null" Then
				AddVerificationInfoToResult "Pass","As expected no rows found"
				LocateColumns = True
				Exit Function
			elseIf totalRow = 0 Then
				AddVerificationInfoToResult "Error","No rows found"
				LocateColumns = False
				Exit Function
			End If

			If currParsedRows> firstRow Then
				currParsedRows = currParsedRows - firstRow +1
			Else
				currParsedRows = 0
			End If
			For i=currParsedRows to lastRow-firstRow
				objSiebList.ActivateRow i
	
				For j=0 to arrCount
					repName=objSiebList.GetColumnRepositoryName (arr(j))
					oDictTemp.Item(arr(j)) = objSiebList.GetCellText (repName, i)
					
				Next
				flag = "N"
	
				For j=0 to arrCount
					flag = "N"
					If instr(oDict.Item(arr(j)),"*")>0 Then
						tempValue = Replace (oDict.Item(arr(j)),"*","")
						If DictionaryTest_G.Exists(tempValue) Then
							
							tempValue = Replace (tempValue,tempValue,DictionaryTest_G(tempValue))
							AddVerificationInfoToResult arr(j),tempValue
						End If
						If DictionaryTempTest_G.Exists(tempValue) Then
							
							tempValue = Replace (tempValue,tempValue,DictionaryTempTest_G(tempValue))
							AddVerificationInfoToResult arr(j),tempValue
						End If
						If instr(1,oDictTemp.Item(arr(j)),tempValue,1) >0 Then
							flag = "Y"
						else
							flag = "N"
							Exit For
						End If
					else
						tempValue = oDict.Item(arr(j))
						If DictionaryTest_G.Exists(tempValue) Then
							If Not(isEmpty(DictionaryTest_G(tempValue))) Then
								tempValue = Replace (tempValue,tempValue,DictionaryTest_G(tempValue))
								AddVerificationInfoToResult arr(j),tempValue
							End If
							
						End If

						If DictionaryTempTest_G.Exists(tempValue) Then
							If Not(isEmpty(DictionaryTempTest_G(tempValue))) Then
								tempValue = Replace (tempValue,tempValue,DictionaryTempTest_G(tempValue))
								AddVerificationInfoToResult arr(j),tempValue
							End If							
						End If
						tempValue = Replace (tempValue,"\","")
						If Ucase(oDictTemp.Item(arr(j))) = Ucase(tempValue) Then
							flag = "Y"
						else
							flag = "N"
							Exit For
						End If
					End If
				Next
	
				If flag = "Y" Then
					If Cstr(actIndex)=Cstr(index) Then
						AddVerificationInfoToResult "Info","Line found"
						LocateColumns = True
						Exit Function
					else
						actIndex = actIndex + 1
					End If
				End If
	
			Next
			parsedRows = parsedRows + i
			currParsedRows = parsedRows
			prevRecord = currRecord
			On error resume next
			objSiebList.NextRowSet
			err.clear
			currRecord = obj.RecordCounter
	
		Loop Until (currRecord=prevRecord)
		LocateColumns = "False-Row Not Exist"
	End If
End Function



'THIS   WAS  DONE  FOR  MENULINEITEM_FN DNDD
Public Function LocateColumns2(obj,sLocateCol,sLocateColValue,index)
                parsedRows=0
                currParsedRows=0
                If instr(sLocateCol,"^") >0  Then
                                LocateColumns2 = LocateAndSelectMultiColumns(obj,sLocateCol,sLocateColValue,index)
                elseIf instr(sLocateColValue,"All|") >0  Then
                                LocateColumns2 = LocateColumnsAll(obj,sLocateCol,sLocateColValue,index)
                else
                
                                Set oDict = CreateObject("Scripting.Dictionary")

                                If instr(sLocateColValue,"Promotion")>0 Then
                                                sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
                                End If
                                arr1 = Split (sLocateCol,"|")
                                arr2 = Split (sLocateColValue,"|")
                                arrCnt1 = Ubound (arr1)
                                For k=0 to arrCnt1
                                                oDict.Add arr1(k),arr2(k)               
                                Next
                
                                Set oDictTemp = CreateObject("Scripting.Dictionary")
                
                                Set objSiebList = obj.SiebList("micclass:=SiebList","repositoryname:=SiebList")
                
                                cnt = objSiebList.ColumnsCount
                                If cnt = 0  Then
                                                AddVerificationInfoToResult "Error","No columns found"
                                                LocateColumns2 = False
                                                Exit Function
                                End If
                
'                               For i=0 to cnt-1
'                                               repName= objSiebList.GetColumnRepositoryNameByIndex(i)
'                                               uiName = objSiebList.GetColumnUIName(repName)
'                                               oDictTemp.Add uiName, i
'                               
'               '                               print uiName
'                               Next
                                arr =  oDict.Keys
                                arrCount= Ubound(arr)
                                For j=0 to arrCount
'                                               If oDictTemp.Exists (arr(j)) = False Then
'                                                               AddVerificationInfoToResult "Error",key1 & "column not found"
'                                                               LocateColumns = False
'                                                               Exit Function
'               
'                                               End If
                                                oDictTemp (arr(j)) = j
                                Next
                
                                currRecord = obj.RecordCounter
                                actIndex = 0

                                Do
                
                                                getRowCount currRecord,firstRow,lastRow,totalRow
                                                If totalRow = 0 and sLocateColValue="Null" Then
                                                                AddVerificationInfoToResult "Pass","As expected no rows found"
                                                                LocateColumns2 = True
                                                                Exit Function
                                                elseIf totalRow = 0 Then
                                                                AddVerificationInfoToResult "Error","No rows found"
                                                                LocateColumns2 = False
                                                                Exit Function
                                                End If

                                                If currParsedRows> firstRow Then
                                                                currParsedRows = currParsedRows - firstRow +1
                                                Else
                                                                currParsedRows = 0
                                                End If
                                                objSiebList.ActivateRow 0
                                                For i=currParsedRows to lastRow-firstRow
                                                                objSiebList.ActivateRow i
                
                                                                For j=0 to arrCount
                                                                                repName=objSiebList.GetColumnRepositoryName (arr(j))
                                                                                oDictTemp.Item(arr(j)) = objSiebList.GetCellText (repName, i)
                                                                                
                                                                Next
                                                                flag = "N"
                
                                                                For j=0 to arrCount
                                                                                flag = "N"
                                                                                If instr(oDict.Item(arr(j)),"*")>0 Then
                                                                                                tempValue = Replace (oDict.Item(arr(j)),"*","")
                                                                                                If DictionaryTest_G.Exists(tempValue) Then
                                                                                                                
                                                                                                                tempValue = Replace (tempValue,tempValue,DictionaryTest_G(tempValue))
                                                                                                                AddVerificationInfoToResult arr(j),tempValue
                                                                                                End If
                                                                                                If DictionaryTempTest_G.Exists(tempValue) Then
                                                                                                                
                                                                                                                tempValue = Replace (tempValue,tempValue,DictionaryTempTest_G(tempValue))
                                                                                                                AddVerificationInfoToResult arr(j),tempValue
                                                                                                End If
                                                                                                If instr(1,oDictTemp.Item(arr(j)),tempValue,1) >0 Then
                                                                                                                flag = "Y"
                                                                                                else
                                                                                                                flag = "N"
                                                                                                                Exit For
                                                                                                End If
                                                                                else
                                                                                                tempValue = oDict.Item(arr(j))
                                                                                                If DictionaryTest_G.Exists(tempValue) Then
                                                                                                                If Not(isEmpty(DictionaryTest_G(tempValue))) Then
                                                                                                                                tempValue = Replace (tempValue,tempValue,DictionaryTest_G(tempValue))
                                                                                                                                AddVerificationInfoToResult arr(j),tempValue
                                                                                                                End If
                                                                                                                
                                                                                                End If

                                                                                                If DictionaryTempTest_G.Exists(tempValue) Then
                                                                                                                If Not(isEmpty(DictionaryTempTest_G(tempValue))) Then
                                                                                                                                tempValue = Replace (tempValue,tempValue,DictionaryTempTest_G(tempValue))
                                                                                                                                AddVerificationInfoToResult arr(j),tempValue
                                                                                                                End If                                                                                                    
                                                                                                End If
                                                                                                tempValue = Replace (tempValue,"\","")
                                                                                                If Ucase(oDictTemp.Item(arr(j))) = Ucase(tempValue) Then
                                                                                                                flag = "Y"
                                                                                                else
                                                                                                                flag = "N"
                                                                                                                Exit For
                                                                                                End If
                                                                                End If
                                                                Next
                
                                                                If flag = "Y" Then
                                                                                If Cstr(actIndex)=Cstr(index) Then
                                                                                                AddVerificationInfoToResult "Info","Line found"
                                                                                                LocateColumns2 = True
                                                                                                Exit Function
                                                                                else
                                                                                                actIndex = actIndex + 1
                                                                                End If
                                                                End If
                
                                                Next
                                                parsedRows = parsedRows + i
                                                currParsedRows = parsedRows
                                                prevRecord = currRecord
                                                On error resume next
                                                objSiebList.NextRowSet
                                                err.clear
                                                currRecord = obj.RecordCounter
                
                                Loop Until (currRecord=prevRecord)
                                LocateColumns2 = "False-Row Not Exist"
                End If
End Function

'Public Function LocateColumns(obj,sLocateCol,sLocateColValue,index)
'	parsedRows=0
'	currParsedRows=0
'	If instr(sLocateCol,"^") >0  Then
'		LocateColumns = LocateAndSelectMultiColumns(obj,sLocateCol,sLocateColValue,index)
'	elseIf instr(sLocateColValue,"All|") >0  Then
'		LocateColumns = LocateColumnsAll(obj,sLocateCol,sLocateColValue,index)
'	else
'	
'		Set oDict = CreateObject("Scripting.Dictionary")
'	
'		If instr(sLocateColValue,"Promotion")>0 Then
'			sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
'		End If
'		arr1 = Split (sLocateCol,"|")
'		arr2 = Split (sLocateColValue,"|")
'		For k=0 to Ubound (arr1)
'			oDict.Add arr1(k),arr2(k)	
'		Next
'	
'		Set oDictTemp = CreateObject("Scripting.Dictionary")
'	
'		Set objSiebList = obj.SiebList("micclass:=SiebList","repositoryname:=SiebList")
'	
'		cnt = objSiebList.ColumnsCount
'		If cnt = 0  Then
'			AddVerificationInfoToResult "Error","No columns found"
'			LocateColumns = False
'			Exit Function
'		End If
'	
'		For i=0 to cnt-1
'			repName= objSiebList.GetColumnRepositoryNameByIndex(i)
'			uiName = objSiebList.GetColumnUIName(repName)
'			oDictTemp.Add uiName, i
'		
'	'		print uiName
'		Next
'		arr =  oDict.Keys
'		arrCount= Ubound(arr)
'		For j=0 to arrCount
'			If oDictTemp.Exists (arr(j)) = False Then
'				AddVerificationInfoToResult "Error",key1 & "column not found"
'				LocateColumns = False
'				Exit Function
'	
'			End If
'		Next
'	
'		currRecord = obj.RecordCounter
'		actIndex = 0
'		Do
'	
'			getRowCount currRecord,firstRow,lastRow,totalRow
'			If totalRow = 0 and sLocateColValue="Null" Then
'				AddVerificationInfoToResult "Pass","As expected no rows found"
'				LocateColumns = True
'				Exit Function
'			elseIf totalRow = 0 Then
'				AddVerificationInfoToResult "Error","No rows found"
'				LocateColumns = False
'				Exit Function
'			End If
'
'			If currParsedRows> firstRow Then
'				currParsedRows = currParsedRows - firstRow +1
'			Else
'				currParsedRows = 0
'			End If
'			For i=currParsedRows to lastRow-firstRow
'				objSiebList.ActivateRow i
'	
'				For j=0 to arrCount
'					repName=objSiebList.GetColumnRepositoryName (arr(j))
'					oDictTemp.Item(arr(j)) = objSiebList.GetCellText (repName, i)
'					
'				Next
'				flag = "N"
'	
'				For j=0 to arrCount
'					flag = "N"
'					If instr(oDict.Item(arr(j)),"*")>0 Then
'						tempValue = Replace (oDict.Item(arr(j)),"*","")
'						If instr(1,oDictTemp.Item(arr(j)),tempValue,1) >0 Then
'							flag = "Y"
'						else
'							flag = "N"
'							Exit For
'						End If
'					else
'						If Ucase(oDictTemp.Item(arr(j))) = Ucase(oDict.Item(arr(j))) Then
'							flag = "Y"
'						else
'							flag = "N"
'							Exit For
'						End If
'					End If
'				Next
'	
'				If flag = "Y" Then
'					If Cstr(actIndex)=Cstr(index) Then
'						AddVerificationInfoToResult "Info","Line found"
'						LocateColumns = True
'						Exit Function
'					else
'						actIndex = actIndex + 1
'					End If
'				End If
'	
'			Next
'			parsedRows = parsedRows + i
'			currParsedRows = parsedRows
'			prevRecord = currRecord
'			On error resume next
'			objSiebList.NextRowSet
'			err.clear
'			currRecord = obj.RecordCounter
'	
'		Loop Until (currRecord=prevRecord)
'		LocateColumns = "False-Row Not Exist"
'	End If
'End Function


Public Function LocateColumnsAll(obj,sLocateCol,sLocateColValue,index)
	parsedRows=0
	currParsedRows=0

	
		Set oDict = CreateObject("Scripting.Dictionary")
	
		If instr(sLocateColValue,"Promotion")>0 Then
			sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
		End If
		arr1 = Split (sLocateCol,"|")
		arr2 = Split (sLocateColValue,"|")
		For k=0 to Ubound (arr1)
			oDict.Add arr1(k),arr2(k)
	
		Next
	
		Set oDictTemp = CreateObject("Scripting.Dictionary")
	
		Set objSiebList = obj.SiebList("micclass:=SiebList","repositoryname:=SiebList")
	
		cnt = objSiebList.ColumnsCount
		If cnt = 0  Then
			AddVerificationInfoToResult "Error","No columns found"
			LocateColumnsAll = False
			Exit Function
		End If
	
		For i=0 to cnt-1
			repName= objSiebList.GetColumnRepositoryNameByIndex(i)
			uiName = objSiebList.GetColumnUIName(repName)
			oDictTemp.Add uiName, i
		
	'		print uiName
		Next
		arr =  oDict.Keys
		arrCount= Ubound(arr)
		For j=0 to arrCount
			If oDictTemp.Exists (arr(j)) = False Then
				AddVerificationInfoToResult "Error",key1 & "column not found"
				LocateColumnsAll = False
				Exit Function
	
			End If
		Next
	
		currRecord = obj.RecordCounter
		actIndex = 0
		Do
	
			getRowCount currRecord,firstRow,lastRow,totalRow
            If totalRow = 0 Then
				AddVerificationInfoToResult "Error","No rows found"
				LocateColumnsAll = False
				Exit Function
			End If

			If currParsedRows> firstRow Then
				currParsedRows = currParsedRows - firstRow +1
			Else
				currParsedRows = 0
			End If
			For i=currParsedRows to lastRow-firstRow
				objSiebList.ActivateRow i
	
				For j=0 to arrCount
					repName=objSiebList.GetColumnRepositoryName (arr(j))
					oDictTemp.Item(arr(j)) = objSiebList.GetCellText (repName, i)
					
				Next
				flag = "N"
	
				For j=0 to arrCount
					flag = "N"

					If Ucase(oDict.Item(arr(j))) = "ALL" Then
						flag = "Y"
					elseIf Ucase(oDictTemp.Item(arr(j))) = Ucase(oDict.Item(arr(j))) Then
						flag = "Y"
						AddVerificationInfoToResult "Pass","Actual for "& oDictTemp.Item(arr(0)) & " is " & oDictTemp.Item(arr(j)) & " as expected"
					else
						flag = "N"
						AddVerificationInfoToResult "Fail","Expected for "& oDictTemp.Item(arr(0)) & " is " & oDict.Item(arr(j)) & "but expected is " & oDictTemp.Item(arr(j)) 
						LocateColumnsAll = False
						Exit Function
						
					End If
				Next
	
			Next
			parsedRows = parsedRows + i
			currParsedRows = parsedRows
			prevRecord = currRecord
			On error resume next
			objSiebList.NextRowSet
			err.clear
			currRecord = obj.RecordCounter
	
		Loop Until (currRecord=prevRecord)
		LocateColumnsAll = True

End Function



Public Function LocateAndSelectMultiColumns(obj,sLocateCols,sLocateColValues,indexes)
	parsedRows=0
	currParsedRows=0

	arr3 =  Split (sLocateCols,"^")
	arr4 =  Split (sLocateColValues,"^")
	arr5 =  Split (indexes,"^")
	Set objSiebList = obj.SiebList("micclass:=SiebList","repositoryname:=SiebList")
	cnt = objSiebList.ColumnsCount
	If cnt = 0  Then
		AddVerificationInfoToResult "Error","No columns found"
		LocateAndSelectMultiColumns = False
		Exit Function
	End If

	Set oDictTemp = CreateObject("Scripting.Dictionary")

	For i=0 to cnt-1
		repName= objSiebList.GetColumnRepositoryNameByIndex(i)
		uiName = objSiebList.GetColumnUIName(repName)
		oDictTemp.Add uiName, i
	
'		print uiName
	Next

	objSiebList.ActivateRow 0
	actIndex = 0
	Dim flag
	For p=0 to Ubound (arr3)
		flag = "N"
		sLocateCol = arr3(p)
		sLocateColValue = arr4(p)
		index = cint (arr5(p))
		
		Set oDict = CreateObject("Scripting.Dictionary")
	
		If instr(sLocateColValue,"Promotion")>0 Then
			sLocateColValue = Replace(sLocateColValue,"Promotion",DictionaryTest_G.Item("ProductName"))
		End If
		arr1 = Split (sLocateCol,"|")
		arr2 = Split (sLocateColValue,"|")
		For k=0 to Ubound (arr1)
			oDict.Add arr1(k),arr2(k)
	
		Next
	
		arr =  oDict.Keys
		arrCount= Ubound(arr)
		For j=0 to arrCount
			If oDictTemp.Exists (arr(j)) = False Then
				AddVerificationInfoToResult "Error",key1 & "column not found"
				LocateColumns = False
				Exit Function
	
			End If
		Next
	
	
		currRecord = obj.RecordCounter
	
	
		Do
	
			getRowCount currRecord,firstRow,lastRow,totalRow
			If totalRow = 0 Then
				AddVerificationInfoToResult "Error","No rows found"
				LocateAndSelectMultiColumns = False
				Exit Function
			End If
			If currParsedRows> firstRow Then
				currParsedRows = currParsedRows - firstRow +1
			Else
				currParsedRows = 0
			End If
			For i=currParsedRows to lastRow-firstRow
				
	
				For j=0 to arrCount
					repName=objSiebList.GetColumnRepositoryName (arr(j))
					oDictTemp.Item(arr(j)) = objSiebList.GetCellText (repName, i)
					
				Next
				
	
				For j=0 to arrCount
					flag = "N"
'					If Ucase(oDictTemp.Item(arr(j))) = Ucase(oDict.Item(arr(j))) Then
'						flag = "Y"
'						
'					else
'						flag = "N"
'						Exit For
'					End If
					tempValue = oDict.Item(arr(j))
						If DictionaryTest_G.Exists(tempValue) Then
							If Not(isEmpty(DictionaryTest_G(tempValue))) Then
								tempValue = Replace (tempValue,tempValue,DictionaryTest_G(tempValue))
								AddVerificationInfoToResult arr(j),tempValue
							End If
							
						End If

						If DictionaryTempTest_G.Exists(tempValue) Then
							If Not(isEmpty(DictionaryTempTest_G(tempValue))) Then
								tempValue = Replace (tempValue,tempValue,DictionaryTempTest_G(tempValue))
								AddVerificationInfoToResult arr(j),tempValue
							End If							
						End If
						tempValue = Replace (tempValue,"\","")
						If Ucase(oDictTemp.Item(arr(j))) = Ucase(tempValue) Then
							flag = "Y"
						else
							flag = "N"
							Exit For
						End If
				Next
	
				If flag = "Y" Then
					If actIndex=index Then
						If p = 0 Then
							objSiebList.ActivateRow i
						else
							objSiebList.SelectRow i,"Ctrl"
						End If
						AddVerificationInfoToResult "Info","Line found"
						Exit Do
					else
						flag = "N"
						actIndex = actIndex + 1
					End If
				End If
	
			Next
			parsedRows = parsedRows + i
			currParsedRows = parsedRows
			prevRecord = currRecord
			On error resume next
			objSiebList.NextRowSet
			err.clear
			currRecord = obj.RecordCounter
	
		Loop Until (currRecord=prevRecord)
		If flag = "N" Then
			AddVerificationInfoToResult "Fail","Row not found"
			Exit Function
		End If
	Next

	If flag = "Y" Then
		LocateAndSelectMultiColumns = True
	else 
		LocateAndSelectMultiColumns = False
	End If

End Function


Public function getRowCount(str,firstRow,lastRow,totalRow)
'   If Right(str,1)="+" Then
'        getRowCount=-1
	If str="No Records" Then
		totalRow = 0
'        getRowCount=0
   else
		rowCount=Split(str,"of ")(1)
		totalRow=Cint(rowCount)
		firstRow=Cint(Split(Split(str,"of ")(0)," - ")(0))
		lastRow=Cint(Split(Split(str,"of ")(0)," - ")(1))
   End If

End Function





