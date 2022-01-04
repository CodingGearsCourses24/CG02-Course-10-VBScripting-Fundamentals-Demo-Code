' WSCRIPT & WSCRIPT are exactly the same except one minor difference
' WSCRIPT.EXE ==> Windows Apps
' CSCRIPT.EXE ==> Console Apps
' WSCRIPT.EXE & CSCRIPT.EXE are part of Windows Script Host (WSH)
' WSCRIPT.EXE vs WSCRIPT
'       WSCRIPT.EXE --> Executable for WSH that interprets VBScript lang
'       WSCRIPT --> An object in Core Object Model in WSH

MsgBox("1 Hello World!")
WScript.Echo "2 Hello World!"

' Sleep for 5 seconds
WScript.Sleep 5000
WScript.Echo "3 Hello World!"
'WScript.Quit
WScript.Echo "4 Hello World!"
