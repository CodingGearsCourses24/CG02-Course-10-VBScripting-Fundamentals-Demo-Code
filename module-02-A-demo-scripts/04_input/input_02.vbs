'=========================================================================
' Using Input --> Getting input from the user
'       Cint    
'=========================================================================

' Variables & Contants
Const  SITE_TITLE = ">> CodingGears.io" 
Dim input1, input2, sum, name

'Getting the input from the user
name = InputBox("Enter your name: ") 
input1 = InputBox("Enter the first number: ") 
input2 = InputBox("Enter the second number: ") 

' Processing
sum = CInt(input1) + CInt(input2)

'Displaying a message
MsgBox "Hello : " & name & " ! ! ! " , 0, SITE_TITLE

MsgBox "The sum of the two numbers : " & sum, 0, SITE_TITLE
