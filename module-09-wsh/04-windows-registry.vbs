'=================================================================================
' Windows Registry
'=================================================================================
' HKEY_CURRENT_USER     --> HKCU
' HKEY_LOCAL_MACHINE    --> HKLM
' HKEY_CLASSES_ROOT     --> HKCR
' HKEY_USERS            --> HKEY_USERS
' HKEY_CURRENT_CONFIG   --> HKEY_CURRENT_CONFIG

Dim oShell, tmp

CreateObjects
'ReadRegistry
'WriteToRegistry
DeleteFromRegistry
CloseObjects

' Create Objects
Sub CreateObjects
    Set oShell = WScript.CreateObject("WScript.Shell")
End Sub

' Write
Sub WriteToRegistry
    oShell.RegWrite "HKCU\Software\CodingGears\CourseName", "VBScripting Fundamentals - WSH Magic", "REG_SZ"
    oShell.RegWrite "HKCU\Software\CodingGears\Website1", "www.GlobalETraining.com", "REG_SZ"
    oShell.RegWrite "HKCU\Software\CodingGears\Website2", "www.CodingGears.io", "REG_SZ"
End Sub

' Read
Sub ReadRegistry
    WScript.Echo oShell.RegRead("HKCU\SOFTWARE\CodingGears\CourseName")
    WScript.Echo oShell.RegRead("HKCU\SOFTWARE\CodingGears\Website1")
    WScript.Echo oShell.RegRead("HKCU\SOFTWARE\CodingGears\Website2")
End Sub

' Delete
Sub DeleteFromRegistry
    oShell.RegDelete "HKCU\Software\CodingGears\CourseName"
    oShell.RegDelete "HKCU\Software\CodingGears\Website1"
    oShell.RegDelete "HKCU\Software\CodingGears\Website2"
End Sub

' Close
Sub CloseObjects
    Set oShell = nothing
End Sub