'=========================================================================
' Loops - FOR
'=========================================================================

option explicit

' Variables
Dim a

' Loop
For  a = 0 To 5
	WScript.Echo a & " : VBScriting is fun!"
Next

WScript.Echo "--------------------"

For  a = 0 To 54 Step 8
	WScript.Echo a & " : VBScriting is fun!"
Next

WScript.Echo "********************"

For  a = 50 To 10 Step -5
	WScript.Echo a & " : VBScriting is fun!"
Next