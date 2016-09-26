set oShell = WScript.CreateObject("WScript.Shell")
oShell.Run "STW.exe -dsn "SilkTestInmetrics" -username "flavio" -password "flavio" -Project "Cotador - RIC - Residencial" -script "MAIN_HOMLOGACAO" -verbose"
'WScript.Sleep 10000
oShell.AppActivate ""
'MsgBox ScriptEngineBuildVersion 'This line shows which version is the Script Engine in XP