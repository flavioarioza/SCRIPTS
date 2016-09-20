If Wscript.Arguments.Count = 0 Then
    strComputer = inputbox("Enter a computer name to Restart","Enter computer name")
    if strComputer = "" then wscript.quit
Else
    strCOmputer = Wscript.Arguments.Item(0)

End If

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Shutdown)}!\\" & _
        strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
    objOperatingSystem.Reboot()
Next