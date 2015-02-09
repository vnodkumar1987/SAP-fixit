Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

'Copy the files if they don't exist, rename them before copying if they exist
'Copying saplogon.ini
If (FSO.FileExists("C:\windows\saplogon.ini")) Then
	'FSO.MoveFile "C:\windows\saplogon.ini", "C:\windows\saplogon.ini.old"
	FSO.DeleteFile "C:\Windows\saplogon.ini"
	FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\saplogon.ini", "C:\windows\saplogon.ini"
	Msgbox "Renamed & copied"
Else
	FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\saplogon.ini", "C:\windows\saplogon.ini"
	Msgbox "Copied"
End If

'***************************************************************

'copying services // double if loop

If (FSO.FileExists("C:\Windows\System32\drivers\etc\services.old")) Then
	FSO.DeleteFile "C:\Windows\System32\drivers\etc\services.old"
	FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\services", "C:\windows\System32\drivers\etc\services"
Else
		If (FSO.FileExists("C:\windows\System32\drivers\etc\services")) Then
			FSO.MoveFile "C:\windows\System32\drivers\etc\services", "C:\windows\System32\drivers\etc\services.old"
			FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\services", "C:\windows\System32\drivers\etc\services"
			Msgbox "Renamed & copied"
		Else
			FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\services", "C:\windows\System32\drivers\etc\services"
			Msgbox "Copied"
		End If
End If
Set FSO = nothing

'*******************************************************************

'copying sapmsg.ini

If (FSO.FileExists("C:\windows\sapmsg.ini")) Then
	'FSO.MoveFile "C:\windows\sapmsg.ini", "C:\windows\sapmsg.ini.old"
	FSO.DeleteFile "C:\Windows\sapmsg.ini"
	FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\sapmsg.ini", "C:\windows\sapmsg.ini"
	Msgbox "Renamed & copied"
Else
	FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\sapmsg.ini", "C:\windows\sapmsg.ini"
	Msgbox "Copied"
End If

Set FSO = nothing

'***************************************************************

'copying saplogon.ini to appdata
Dim LocalAppData

'Set WshShell = CreateObject("Wscript.Shell")
'Set FSO = CreateObject("Scripting.FileSystemObject")

LocalAppData = WshShell.ExpandEnvironmentStrings("%APPDATA%")

If (FSO.FileExists(LocalAppData & "\SAP\Common\saplogon.ini")) Then
	FSO.DeleteFile(LocalAppData & "\SAP\Common\saplogon.ini")
	FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\saplogon.ini", LocalAppData & "\SAP\Common\saplogon.ini"
	'Msgbox "Renamed & copied"
Else
	FSO.CopyFile "\\indshare\apps\NonAdobe\Win\SAP\SAPGUI_720\CustomerFiles\saplogon.ini", LocalAppData & "\SAP\Common\saplogon.ini"
	'Msgbox "Copied"
End If

'***************************************************************
