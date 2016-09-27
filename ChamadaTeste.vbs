set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run "STW.exe -dsn SilkTestInmetrics -username flavio -password flavio -Project Cotador_RIC_Residencial -script MAIN_HOMOLOGACAO -verbose"
'WScript.Sleep 5000
oShell.AppActivate ""


