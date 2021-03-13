'=========================================================================
' Loop: Do Until
'=========================================================================

Option Explicit

' Variables
Dim a

a = 1

'Processing
Do Until a = 5 'Exit on True
	WScript.Echo "a = " & a 
	a = a + 1
	WScript.Echo "The value of a after incrementing : " & a
Loop

WScript.Echo "1: Out side the loop!"
