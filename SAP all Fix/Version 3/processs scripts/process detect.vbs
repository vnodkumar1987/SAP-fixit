Option Explicit
Dim objWMIService, objProcess, colProcess
Dim strComputer, strList

strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _ 
& strComputer & "\root\cimv2") 

Set colProcess = objWMIService.ExecQuery _
("Select * from Win32_Process")

For Each objProcess in colProcess
    strList = strList & vbCr & _
    objProcess.Name
Next

WSCript.Echo strList
WScript.Quit