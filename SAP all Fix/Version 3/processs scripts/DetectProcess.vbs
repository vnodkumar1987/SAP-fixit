set objWMIService = GetObject ("winmgmts:")
foundProc = False
procName="notepad.exe"
procNameFriend="Word"

for each Process in objWMIService.InstancesOf("Win32_Process")
If StrComp(Process.Name,procName,vbTextCompare)=0 then
foundProc=true
End If
Next
'If foundProc = True Then
' WScript.Echo "Found Process"
'End If

Do While foundProc = true
	msgbox "fnud"
Loop