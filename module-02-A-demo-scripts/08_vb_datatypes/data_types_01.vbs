'=========================================================================
' VBScript has only one data type called a Variant
' Empty, Null, Boolean, Byte, Integer, Currency
' Long, Single, Double, Date (Time), String
' Object, Error
'=========================================================================

Dim strName, arrNames, dtStartDate, result, city, count, temperature

strName = "Coding"
arrNames = Array("Peter", "Mary", "Kelly")
dtStartDate = Now
intCount = 100
temperature = 25.25
cAmount1 = 123.3
cAmount2 = 555.635

result = VarType(dtStartDate)
WScript.Echo "Result : ==> " & result

If VarType(dtStartDate) = vbDate Then
    WScript.Echo "It's a date type"
End If
