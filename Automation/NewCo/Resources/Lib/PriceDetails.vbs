'#################################################################################################
' 	Function Name : PriceCalculations_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function PriceCalculations_fn()

   	Dim adoData	  
    strSQL = "SELECT * FROM [PriceCalculations$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PriceDetails.xls",strSQL)

	sType = adoData("Type") & ""
	sDuration_Quantity = adoData("Duration_Quantity") & ""
	If DictionaryTest_G.Exists(sDuration_Quantity) Then
		sDuration_Quantity = Replace (sDuration_Quantity,sDuration_Quantity,DictionaryTest_G(sDuration_Quantity))
	End If
	sCriteria = adoData("Criteria") & ""
	StdPrice = DictionaryTest_G("StdPrice")
	

	If  sType = "Voice" Then
		Beat = DictionaryTest_G("Beat")
		TotalMins = sDuration_Quantity \ Beat
		If sDuration_Quantity mod Beat > 0 Then
			TotalMins = TotalMins + 1
		End If
		TotalPrice = StdPrice * TotalMins

	elseIf  sType = "SMS" Then

		TotalPrice = StdPrice * sDuration_Quantity

	elseif sType = "InBundle" Then
		TotalPrice = 0		
	elseif sType = "Fixed" Then
		TotalPrice = sDuration_Quantity
	End If

	If sCriteria = "Traveller" Then
		fee = DictionaryTest_G("Fee")
		TotalPrice = TotalPrice + fee
	End If

	DictionaryTest_G("TotalPrice") = TotalPrice

End Function
