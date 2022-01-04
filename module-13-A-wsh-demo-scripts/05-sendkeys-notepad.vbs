'=================================================================================
' Sending Keys Strokes
'=================================================================================
'
Dim oShell
Const time_in_milliseconds = 1000

CreateObjects
Wait time_in_milliseconds
OpenNotepad
SendText
CloseNotepad
CloseObjects

'=================================================================================
' Create Objects
'=================================================================================
Sub CreateObjects
    Set oShell = WScript.CreateObject("WScript.Shell")
End Sub

'=================================================================================
' Wait
'=================================================================================
Sub Wait(time_in_milliseconds)
    WScript.Sleep time_in_milliseconds 
    WScript.Echo "Waited for " & time_in_milliseconds & " milli-seconds!"
End Sub

'=================================================================================
' Open Notepad
'=================================================================================
Sub OpenNotepad
    oShell.Run "notepad.exe"
End Sub

'=================================================================================
' Send Text
'=================================================================================
Sub SendText
    WScript.Sleep 500
    oShell.AppActivate "untitled - Notepad" 
    oShell.SendKeys "I am learning VBScripting Fundamentals!"
    oShell.SendKeys "{ENTER}"
    oShell.SendKeys "{F5}"
    WScript.Sleep 500
End Sub

'=================================================================================
' Close Notepad
'=================================================================================
Sub CloseNotepad
    oShell.SendKeys "%F"  'Alt F
    oShell.SendKeys "x"
    oShell.SendKeys "{ENTER}"
    oShell.SendKeys "D:\zTemp1\sample.txt"
    oShell.SendKeys "{ENTER}"
End Sub

'=================================================================================
' Close Objects
'=================================================================================
Sub CloseObjects
    Set oShell = nothing
End Sub