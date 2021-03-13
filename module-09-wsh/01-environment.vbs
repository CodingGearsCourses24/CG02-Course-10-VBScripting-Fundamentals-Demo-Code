'=================================================================================
' WshShell --> WshEnvironment
'=================================================================================


Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
' The WshEnvironment object is a collection of environment variables 
' that is returned by the WshShell object's Environment property.
Set WshEnv = WshShell.Environment("SYSTEM")


Welcome
'RunEnvironment
GetEnvironmentVariables
CloseObjects


'=================================================================================
' Welcome
'=================================================================================
Sub Welcome
    Wscript.Echo "Welcome to CodingGears.io"
    WScript.Echo WshShell.CurrentDirectory
    Wscript.echo "----------------------------------------------------"

End Sub

'=================================================================================
' Environment
'=================================================================================
Sub RunEnvironment
    WScript.Echo "Processors : " & WshEnv("NUMBER_OF_PROCESSORS")
    WScript.Echo "Processor Architecture : " & WshEnv("PROCESSOR_ARCHITECTURE")
    WScript.Echo "OS : " & WshEnv("OS")
    'WScript.Echo WshEnv("PATH")
    WScript.Echo "Temp Folder : " & WshEnv("TEMP")
    Wscript.echo "----------------------------------------------------"
    Wscript.echo "Processors : " & WshEnv("NUMBER_OF_PROCESSORS") & vbCRLF &_
                 "Processor: " & WshEnv("PROCESSOR_ARCHITECTURE") & vbCRLF &_
                 "Operating system: " & WshEnv("OS")
End Sub

'=================================================================================
' Access Environment Variables
'=================================================================================
Sub GetEnvironmentVariables
    WScript.Echo "Username : " & WshShell.ExpandEnvironmentStrings("%UserName%")
    WScript.Echo "Computer Name : " & WshShell.ExpandEnvironmentStrings("%ComputerName%")
    WScript.Echo "Course Name : " & WshShell.ExpandEnvironmentStrings("%CourseName%")
    WScript.Echo "Website : " & WshShell.ExpandEnvironmentStrings("%Website%")

End Sub

'=================================================================================
' Closing Objects
'=================================================================================
Sub CloseObjects
    Set WshShell = Nothing
    Set WshEnv = nothing
End Sub