' Using built-in functions - Formatting
' FormatCurrency, FormatDateTime, FormatNumber, FormatPercent
' ------------------------------------------------------------------
number = -12345.6789123
number2 = 12568956.256 

' >>>>> Percentage
' FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
mypercent1 = FormatPercent(45/80) ' Default
mypercent2 = FormatPercent(45/80, 8, vbTrue) ' 8 decimals, leading 0s
mypercent3 = FormatPercent(-45/80, 8, vbTrue, vbTrue) ' Indicates whether or not to place negative values within parentheses
mypercent4 = FormatPercent(-45/0.25, 8, vbTrue, vbTrue, True) ' Indicates whether or not numbers are grouped

' DisplayMessage mypercent4, "MyPercent"

' >>>>> Currency
' FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) 
DisplayMessage FormatCurrency(number2, 2), "2 Decimals"
DisplayMessage FormatCurrency(number, 12, vbTrue, vbTrue), " 12 decimals, leading 0s, use () "
DisplayMessage FormatCurrency(number2, 5, vbTrue, , vbTrue), " 5 decimals, leading 0s, use (), grouping"


Function DisplayMessage(message, id)
	MsgBox id & " : " & message,0,"Welcome"
End Function