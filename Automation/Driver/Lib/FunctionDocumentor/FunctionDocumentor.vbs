	Set fso = CreateObject("Scripting.FileSystemObject")
	ProductName =InputBox("Please Enter Product Name To Document :")

	'---------------------------------------------------------------------------------------
	SourcePath= "C:\QA_Auto_QTP\"&ProductName&"\Resources\Lib"
	DestinationPath = "C:\QA_Auto_QTP\"&ProductName
	'--------------------------------------------------------------------------------------------------------
	If not(fso.FolderExists(SourcePath)) Then
		msgbox "Invalid Product Name"
		Wscript.Quit 
	End If
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fso, f, Counter
	Set Dictionary_G = CreateObject("Scripting.Dictionary")
	Dictionary_G.Add "TotalCount",0
	Set folder = fso.GetFolder(SourcePath)  '"C:\New Folder\Data"
	Set files = folder.Files
	Dim Ac
	Ac = Array("TOC","Description")
	Counter=1
	For a= 0 to 1
		tot=0
		Dictionary_G.Remove "TotalCount"
		Dictionary_G.Add "TotalCount",tot
		For each folderIdx In files
			If Instr(1,folderIdx.Name,".vbs") > 0 then
				Call WriteFile(Counter,Ac(a),folderIdx.Name)
			End if
			Counter=Counter + 1
		Next
	Next			   
	Set DocFile = fso.OpenTextFile(DestinationPath&"\FunctionDocumentation.html",ForAppending, True)
	DocFile.Write "</html>"

	Function WriteFile(i,Action,FileName)
		If i=1 Then
			Set DocFile = fso.CreateTextFile(DestinationPath&"\FunctionDocumentation.html", True)
		Else
			Set DocFile = fso.OpenTextFile(DestinationPath&"\FunctionDocumentation.html",ForAppending, True)
		End If
		Set AllCode = fso.OpenTextFile(SourcePath&"\"&FileName, ForReading,,-1)
		If i=1 Then
			DocFile.Write  "<a name=TOP></a>"
			DocFile.Write "<html><head><font SIZE=5 COLOR=#0000c0><b>Table Of Content</b></font></head><br>"
		End if
		Do While AllCode.AtEndOfStream <> True
			NFoundFlag = 0
			CurLine = AllCode.ReadLine()
			'Check for Name fo the function
			If  instr(1,ucase(CurLine),"FUNCTION NAME" ) > 0 then
				NameArr = Split(CurLine,":")
				If Ubound(NameArr) = 1 Then
					FName = Trim(NameArr(1)) 'Check next 4 lines for Descrition and other
					DescriptionArr = AllCode.ReadLine()
					If  instr(1,UCASE(DescriptionArr),"DESCRIPTION" ) > 0 and instr(1,CurLine,":" ) > 0 then
						FDescription =Trim(Split(DescriptionArr,":")(1))
					End if
					NewLine=AllCode.ReadLine()
					do While not (instr(1,ucase(NewLine),"END FUNCTION") > 0)
							If  (Instr(1,ucase(NewLine),"SELECT * FROM") > 0) And (Instr(1,NewLine,"$") > 0 )Then
								SheetName1 = Split(NewLine,"$")
								SheetName2=Split(SheetName1(0),"[")
								SheetName=SheetName2(1)
							End If
							If  ((Instr(1,UCASE(NewLine),"EXECEXCQRY") > 0) And (Instr(1,NewLine,".xls") > 0 )And (Instr(1,NewLine,"Data\") >0)) Then
								WB1 = Split(NewLine,".xls")
								WB2=Split(WB1(0),"\")
								WBName=WB2(1)&".xls"
							End If
							If  (Instr(1,UCASE(NewLine),Ucase("SetObjRepository")) > 0) Then
								TSR1 = Split(NewLine,"(")
								TSR2=Split(TSR1(1),",")
								TotLen = Len(Trim(TSR2(0)))
								TSR3 = Left(Right(TSR2(0),TotLen-1),TotLen-2)
								TSRName=TSR3&".tsr"
							End If
							If not AllCode.AtEndOfStream <> True Then
							   Exit Do
							Else
								NewLine=AllCode.ReadLine()
							End If
				  Loop

'					CreatedByArr =AllCode.ReadLine()
'					If  instr(1,CreatedByArr,"Created By" ) > 0 and instr(1,CurLine,":" ) > 0 then
'						FCreatedBy = Trim(Split(CreatedByArr,":")(1))
'					End if
'					CreationDateArr = AllCode.ReadLine()
'					If  instr(1,CreationDateArr,"Creation Date" ) > 0 and instr(1,CurLine,":" ) > 0 then
'						FCreationDate = Trim(Split(CreationDateArr,":")(1))
'					End if
					If Action="TOC" Then
						tot1=Dictionary_G.Item("TotalCount")
						tot1=tot1+1
						Dictionary_G.Remove "TotalCount" 
						Dictionary_G.Add "TotalCount",tot1
						DocFile.Write tot1&"&nbsp&nbsp&nbsp&nbsp<a href=#"&FName&">"&FName&"</a><br>"	
					Elseif Action="Description" then
						DocFile.Write "<a name="&FName&"></a>"
						tot1=Dictionary_G.Item("TotalCount")
						tot1=tot1+1
						Dictionary_G.Remove "TotalCount"
						Dictionary_G.Add "TotalCount",tot1
						DocFile.Write "<TABLE BORDER=1 WIDTH=100% CELLPADDING=3 CELLSPACING=0 SUMMARY=><TR BGCOLOR=CCCCFF CLASS=TableHeadingColor><TH ALIGN=left COLSPAN=2><FONT SIZE=4>"
						DocFile.Write tot1&".Function Name&nbsp&nbsp&nbsp"& FName 
						DocFile.Write "</B></FONT>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<a href=#TOP>TOP</a><br>"
						DocFile.Write "</TH></TR></TABLE>"
						DocFile.Write "VBS File Name : "&FileName&"<br>"
						DocFile.Write "Function Description  :"& FDescription &"<br>"
						DocFile.Write "Data File Name :"&WBName&"<br>"
						
'						DocFile.Write "Data File Name :<a href="&DestinationPath&"\Data\"&WBName&"#'"&SheetName&"'!A1>"&WBName&"</a><br>"
						
						DocFile.Write "Data Sheet name :<a href="&DestinationPath&"\Data\"&WBName&"#'"&SheetName&"'!A1 TARGET=_BLANK>"& SheetName &"</a><br>"
						DocFile.Write "Object Repository Used :"& TSRName&"<br>"
						SheetName =""
						WBName=""
						TSRName=""
'						DocFile.Write "Created By  : "& FCreatedBy &"<br>"
'						DocFile.Write "Creation Date  : "& FCreationDate&"<br>"
						DocFile.Write ("<hr align=left width=100% height=2>")
					End if
				Else
					If Action="TOC" Then
						DocFile.Write "<font color=Red>Syntax Error !!</font><br>"       'Write Error in file
					Elseif Action ="Description" then
						DocFile.Write "<font color=Red>Syntax Error !!</font>"       'Write Error in file
						DocFile.Write ("<hr align=left width=1500 height=2>")
					End If
				End If
			End if	
		Loop
		AllCode.Close
		DocFile.Close
	End Function





