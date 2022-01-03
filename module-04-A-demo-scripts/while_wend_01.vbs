'=========================================================================
' Loop: While..Wend
'=========================================================================

option explicit

' Variables
Dim Counter

Counter = 0 

While Counter < 10   ' Loop runs until the condition is true
   WScript.Echo  "Counter: " & Counter
   Counter = Counter + 1   ' Increment Counter.
Wend   

WScript.Echo  
WScript.Echo  "Done: Outside the while-wend loop"