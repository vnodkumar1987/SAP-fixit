strComputer = "." 
strProcName = "notepad.exe" 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
Set colMonitorProcess = objWMIService.ExecNotificationQuery _ 
 ("SELECT * FROM __InstanceCreationEvent WITHIN 1" & _ 
 "WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name='" & _ 
 strProcName & "'")  
'WScript.Echo "Waiting for " & strProcName & " process to be created ..." 
Set objLatestEvent = colMonitorProcess.NextEvent 
Wscript.Echo VbCrLf & "Process Name: " & objLatestEvent.TargetInstance.Name 
'Wscript.Echo "Process ID: " & objLatestEvent.TargetInstance.ProcessId 
'WScript.Echo "Time: " & Now 