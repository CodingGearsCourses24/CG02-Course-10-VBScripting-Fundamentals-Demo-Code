' TypeName, VarType

Option Explicit

Dim strName, arrNames, dtStartDate, result, city, count, temperature

strName = "CodingGears"
arrNames = Array("Peter", "Mary", "Kelly")
dtStartDate = Now
count = 100
temperature = 25.25

result = VarType(temperature)
WScript.Echo "Result : " & result