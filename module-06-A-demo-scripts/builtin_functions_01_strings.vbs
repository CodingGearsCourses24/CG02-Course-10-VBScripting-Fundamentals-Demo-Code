'=========================================================================
' Built-in String Functions: 
' Mid, Len, StrReverse, LCase, UCase
' Left, Right, LTrim, RTrim, Trim, Replace, StrComp, InStr
'=========================================================================

site = "www.CodingGears.com"
message1 = "I am learning VBScripting at CodingGears.com"
message2 = "   CodingGears     "

' Mid
result1 = mid(site, 5, 2)
DisplayMessage result1, "Mid"

' Len
result2 = Len(site)
DisplayMessage result2, "Len"

' StrReverse
result3 = StrReverse(site)
DisplayMessage result3, "StrReverse"

' LCase
result4 = LCase(site)
DisplayMessage result4, "LCase"

' UCase
result5 = UCase(site)
DisplayMessage result5, "UCase"

' Left
result6 = Left(site, 3)
DisplayMessage result6, "Left"

' Right
result7 = Right(site, 6)
DisplayMessage result7, "Right"

' LTrim
result8 = LTrim(message2)
DisplayMessage result8, "LTrim"

' RTrim
result9 = RTrim(site)
DisplayMessage result9, "RTrim"

' Trim
result10 = Trim(site)
DisplayMessage result10, "Trim"

' Replace
result11 = Replace(site, "CodingGears", "GlobalETraining" )
DisplayMessage result11 , "Result: "

' StrComp
result12 = StrComp("CodingGears", "codingGears", 1)
DisplayMessage result12 , "Result: "

' InStr
result13 = InStr(message1, "VBScripting" )
DisplayMessage result13 , "Result: "



Function DisplayMessage(message, id)
	MsgBox id & " :  " & message,0,"> > > > > Welcome < < < < <"
End Function