' Using built-in functions - Formatting
' FormatCurrency, FormatDateTime, FormatNumber, FormatPercent
' ------------------------------------------------------------------
number = -1212345.6789123
number2 = 12568956.256 


' >>>>> Number format
' FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) 
' DisplayMessage FormatNumber(number, 2), "2 decimals"
' DisplayMessage FormatNumber(number, 12, vbTrue), " 12 decimals, leading 0s"
' DisplayMessage FormatNumber(number, 8, vbTrue, vbTrue), " 8 decimals, leading 0s, use()"
' DisplayMessage FormatNumber(number2, 2, vbTrue, vbTrue, vbTrue), " 8 decimals, leading 0s, use(), grouping"

' >>>>> FormatDateTime 
' FormatDateTime(Date, NamedFormat) 
dt1 = FormatDateTime(Date, vbLongDate) 		' weekday, monthname, year
dt2 = FormatDateTime(Date, vbShortDate ) 	' mm/dd/yyyy
dt3 = FormatDateTime(Date, vbLongTime ) 	' hh:mm:ss PM/AM
dt4 = FormatDateTime(Date, vbShortTime ) 	' hh:mm
dt5 = FormatDateTime(Date, vbGeneralDate ) 	' Default mm/dd/yyyy 

DisplayMessage dt1, "vbLongDate"
DisplayMessage dt2, "vbShortDate"
DisplayMessage dt3, "vbLongTime"
DisplayMessage dt4, "vbShortTime"
DisplayMessage dt5, "vbGeneralDate"


Function DisplayMessage(message, id)
	MsgBox id & " : " & message,0,"Welcome"
End Function
