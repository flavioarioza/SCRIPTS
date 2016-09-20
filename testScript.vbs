
MsgBox "You entered:" & vbCrLf & TextEntry("Disserte sobre sua opção sexual:", "")




Function TextEntry(strPrompt, strTitle)
Dim fs, web, doc
Dim strFile, strChoice
Dim intChars
Dim dtTime
	On Error Resume Next
	Set web = CreateObject("InternetExplorer.Application")
	If web Is Nothing Then
		TextEntry = ""
		Exit Function
	End If
	web.Width = 350
	web.Height = 250
	web.Offline = True
	web.AddressBar = False
	web.MenuBar = False
	web.StatusBar = False
	web.Silent = True
	web.ToolBar = False
	web.Navigate "about:blank"
	'Wait for the browser to navigate to nowhere
	dtTime = Now
	Do While web.Busy
		'Don't wait more than 5 seconds
		Wscript.Sleep 100
		If (dtTime + 5/24/60/60) < Now Then
			TextEntry = ""
			web.Quit
			Exit Function
		End If
	Loop
	'Wait for a good reference to the browser document
	Set doc = Nothing
	dtTime = Now
	Do Until Not doc Is Nothing
		Wscript.Sleep 100
		Set doc = web.Document
		'Don't wait more than 5 seconds
		If (dtTime + 5/24/60/60) < Now Then
			TextEntry = ""
			web.Quit
			Exit Function
		End If
	Loop
	'Write the HTML form
	doc.Write "<html><head><title>" & strTitle & "</title></head>"
	doc.Write "<body><b>" & strPrompt & "</b><br><form><textarea "
	doc.Write "cols=30 rows=5 name=textentry>"
	doc.Write "</textarea><br><br><input type=button "
	doc.Write "name=submit value=""OK"" onclick='javascript:submit.value=""Done""'>"
	doc.Write "</form></body></html>"
	'Show the form
	web.Visible = True
	'Wait for the user to choose, but fail gracefully if a popup killer.
	Err.Clear
	Do Until doc.Forms(0).elements("submit").Value <> "OK"
		Wscript.Sleep 100
		If doc Is Nothing Then
			TextEntry = ""
			web.Quit
			Exit Function
		End If
		If Err.Number <> 0 Then
			TextEntry = ""
			web.Quit
			Exit Function
		End If
	Loop
	'Retrieve the chosen value
	TextEntry = doc.Forms(0).elements("textentry").Value
	web.Quit
End Function