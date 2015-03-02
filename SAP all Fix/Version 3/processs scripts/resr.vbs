OPTION EXPLICIT

    DIM strComputer, strProcess, strUserName, wshShell

    Set wshShell = WScript.CreateObject( "WScript.Shell" )
    strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
    strComputer = "."
    strProcess = "notepad.exe" 'change this to whatever you are trying to detect

    IF isProcessRunning(strComputer, strProcess, strUserName) THEN
        If MsgBox ("Notepad needs to be closed.", 1) = 1 then
            wscript.Quit(1) 'you need to terminate the process if that's your intention before quitting
        End If
    Else
        msgbox ("Process is not running") 'optional for debug, you can remove this
    END IF

FUNCTION isProcessRunning(ByRef strComputer, ByRef strProcess, ByRef strUserName)

    DIM objWMIService, strWMIQuery, objProcess, strOwner, Response

    strWMIQuery = "SELECT * FROM Win32_Process WHERE NAME = '" & strProcess & "'"

    SET objWMIService = GETOBJECT("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2").ExecQuery(strWMIQuery)

    IF objWMIService.Count > 0 THEN
        msgbox "We have at least ONE instance of Notepad"
        For Each objProcess in objWMIService
            Response = objProcess.GetOwner(strOwner)
            If Response <> 0 Then
                'we didn't get any owner information - maybe not permitted by current user to ask for it
                Wscript.Echo "Could not get owner info for process [" & objProcess.Name & "]" & VBNewLine & "Error: " & Return
            Else 
                Wscript.Echo "Process [" & objProcess.Name & "] is owned by [" & strOwner & "]" 'for debug you can remove it
                if strUserName = strOwner Then
                    msgbox "we have the user who is running notepad"
                    isProcessRunning = TRUE
                Else
                    'do nothing as you only want to detect the current user running it
                    isProcessRunning = FALSE
                End If
            End If
        Next
    ELSE
        msgbox "We have NO instance of Notepad - Username is Irrelevant"
        isProcessRunning = FALSE
    END If

End Function