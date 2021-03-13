'=========================================================================
' Sub Procedure --- does not return a value
' Function Procedure ------- can return a value. 
'=========================================================================

option explicit

' Variables
Dim temp, inCelsius, tempInput, x

Call ConvertTemp

'A Sub procedure -- does not return a value
Sub ConvertTemp
   temp = InputBox("Please enter the temperature in degrees F.", 1)
   MsgBox "The temperature is " & Celsius(temp) & " degrees Celsius.", 64
End Sub

'A Function procedure -- can return a value. 
Function Celsius(fDegrees)
    x = (fDegrees - 32) * 5 / 9
    Celsius = Round(x, 2)
End Function