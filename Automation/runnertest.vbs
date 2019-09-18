On error resume next
Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("User")
objEnv("Module") = WScript.Arguments.Item(0)
objEnv("Environment") = WScript.Arguments.Item(1)
objEnv("TCResult") = ""
'msgbox objEnv("Module")
'msgbox objEnv("Environment")
TCResults=objEnv("TCResult")
arrTCResults=Split(TCResults,";")
WScript.StdOut.WriteLine "Start"
For each arrTCResult in arrTCResults
	'TCName = Split(arrTCResult)(0)
	'Result = Split(arrTCResult)(1)
	WScript.StdOut.WriteLine arrTCResult
Next
WScript.StdOut.WriteLine "End"
TCResults="xfew"
If err.number <> 0 or TCResults = "" Then
	WScript.Quit(2)
else
	WScript.Quit(0)
End If
