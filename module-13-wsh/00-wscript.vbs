' WScript
' The WScript object is the root object 
' of the Windows Script Host object model hierarchy.

WScript.Echo "Hello...!"
' WScript.Sleep 2000 ' 1 Sec


WScript.Echo WScript.Name
WScript.Echo WScript.Path
WScript.Echo WScript.ScriptFullName
WScript.Echo WScript.ScriptName
WScript.Echo WScript.Version
WScript.Echo WScript.BuildVersion

' WScript.Quit

Set owshshell = WScript.CreateObject("WScript.Shell")