'=========================================================================
' Loop: Exiting
'=========================================================================

Option Explicit

' Variables
Dim a

a = 1

Do While a < 25 'Exit when false
	MsgBox "a = " & a 
	WScript.Echo a & " - Welcome to CodingGears.com"
    If a = 5 Then
        WScript.Echo "a is equal to 5"
        WScript.Echo "Exiting..."
        Exit Do
    End If
    a = a + 1
Loop

WScript.Echo "1: Out side the loop!"
