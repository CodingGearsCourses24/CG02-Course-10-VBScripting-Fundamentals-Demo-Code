'=========================================================================
' Using Input --> Getting input from the user
'       CInt & Customizing the input box
'=========================================================================

' Variables & Contants
Const  SITE_TITLE = ">> CodingGears.io" 
Dim input1, input2, sum, name

'Getting the input from the user

'Simple
name = InputBox("Enter your name: ") 

'Adding little customization: Default text
input1 = InputBox("Enter the first number: ", SITE_TITLE, "Enter Input Here" ) 

'More Customization: Move the input box around
input2 = InputBox("Enter the second number: ", SITE_TITLE, "Enter Input Here", 1000, 5000) 

sum = CInt(input1) + CInt(input2)

'Displaying a message
MsgBox "Hello : " & name & " ! ! ! " , 0, SITE_TITLE

MsgBox "The sum of the two numbers : " & sum, 0, SITE_TITLE
