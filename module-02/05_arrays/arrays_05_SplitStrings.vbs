'=========================================================================
' Arrays - Split Function
'=========================================================================

' Error Handling
option explicit

' Variables
Const  SITE_TITLE1 = "www.CodingGears.com" 

Dim arrTitle1
arrTitle1 = Split(SITE_TITLE1, ".")

WScript.Echo arrTitle1(0)
WScript.Echo arrTitle1(1)
WScript.Echo arrTitle1(2)