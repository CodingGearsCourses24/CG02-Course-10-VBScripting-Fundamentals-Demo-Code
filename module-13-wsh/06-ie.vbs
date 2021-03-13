'=================================================================================
' InternetExplorer.Application
'=================================================================================

Dim strMessage

' Open Internet Explorer 
Set objIE = WScript.CreateObject("InternetExplorer.Application")
Set objShell = WScript.CreateObject("WScript.Shell")

objIE.Navigate "http://www.CodingGears.io"
objIE.Toolbar = True
objIE.StatusBar = False
objIE.MenuBar = False

strMessage = "Do you want to launch IE?"

userSelection = objShell.Popup(strMessage, , "Scripting IE", vbYesNo + vbQuestion)

If userSelection = vbYes Then
    objIE.Visible = True
Else
    objIE.Quit
End If

Set objIE = Nothing
Set objShell = Nothing