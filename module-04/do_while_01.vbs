'=========================================================================
' Loop: Do..While
'=========================================================================

Option Explicit

' Variables
Dim a

a = 1

Do While a < 5 'Exit when false
	WScript.Echo "a = " & a 
	a = a + 1
	WScript.Echo "The value of a after incrementing : " & a
Loop

WScript.Echo "1: Out side the loop!"
