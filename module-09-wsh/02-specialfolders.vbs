'=================================================================================
' WshShell --> SpecialFolders
'=================================================================================

Dim owshShell
Set owshShell = WScript.CreateObject("WScript.Shell")

Welcome
RunSpecialFolders
CloseObjects

'=================================================================================
' Welcome
'=================================================================================
Sub Welcome
    Wscript.Echo "Welcome to CodingGears.io"
End Sub

'=================================================================================
'SpecialFolders
'=================================================================================
Sub RunSpecialFolders
    WScript.Echo owshShell.SpecialFolders("Desktop")
    WScript.Echo owshShell.SpecialFolders("MyDocuments")
    WScript.Echo owshShell.SpecialFolders("Startup")
End Sub

'=================================================================================
' Closing Objects
'=================================================================================
Sub CloseObjects
    Set owshShell = Nothing
End Sub