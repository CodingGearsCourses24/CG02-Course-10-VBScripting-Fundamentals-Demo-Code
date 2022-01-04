'=================================================================================
' Run, Popup, Exec
'=================================================================================

Dim owshShell
Set owshShell = WScript.CreateObject("WScript.Shell")

Welcome
'RunExec
RunPopup
'RunRun
CloseObjects


'=================================================================================
' Welcome
'=================================================================================
Sub Welcome
    Wscript.Echo " Welcome to CodingGears.io"
    WScript.Echo "---------------------------"
End Sub

'=================================================================================
'Exec - Runs an application in a child command-shell
'=================================================================================
Sub RunExec
    owshShell.Exec("calc")
End Sub

'=================================================================================
'Popup
'=================================================================================
Sub RunPopup
    btn = owshShell.Popup("Do you feel alright?", 5, "Question:", 4 + 32)

    Select Case btn
        ' Yes button pressed.
        case vbYes
            WScript.Echo "Glad to hear you feel alright."
        ' No button pressed.
        case vbNo
            WScript.Echo "Hope you will feel better soon."
        ' Timed out.
        case -1
            WScript.Echo "Hello.... anyone there???"
    End Select
End Sub

'=================================================================================
' Run - Runs a program in a new process.
'=================================================================================
Sub RunRun
    owshShell.run "cmd"
End Sub

'=================================================================================
' Closing Objects
'=================================================================================
Sub CloseObjects
    Set owshShell = Nothing
End Sub