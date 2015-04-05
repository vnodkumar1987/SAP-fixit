dim objShell
dim objFolder
					
set objShell = CreateObject("shell.application")
set objFolder = objShell.NameSpace("c:\Windows\Temp")
objFolder.CopyHere("\\indshare.corp.adobe.com\apps\NonAdobe\Win\SAP\SAPGUI_720\")

set objShell = nothing
set objFolder = nothing


set WshShell = CreateObject("WScript.shell")
WshShell.run ("C:\Windows\Temp\SAPGUI_720\Setup\NwSapSetup.exe /package=Adobe_V1 /noDlg")