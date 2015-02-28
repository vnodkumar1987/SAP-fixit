Const HKEY_CURRENT_USER = &H80000001
strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Office\14.0\Excel\Options"
strValueName = "OPEN"
strValue = """C:\Program Files (x86)\Common Files\SAP Shared\BW\sapbex.xla"""
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue