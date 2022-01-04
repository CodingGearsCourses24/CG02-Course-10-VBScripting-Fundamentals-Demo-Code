' Assignment Solution

' Thompson - Presisent
' Rooter - Sr. Vice President
' Cooper - Vice President
' Parker - Manager

LastName = InputBox("Please enter your last name: ") 

Company ="ABC Inc."

Select Case LastName
  Case "Thompson"
    MsgBox "You are the President of " & Company
  Case "Rooter"
    MsgBox "You are the Sr. Vice President of " & Company
  Case "Cooper"
    MsgBox "You are the Vice President of " & Company
  Case "Parker"
    MsgBox "You are the Manager of " & Company
  Case else
    MsgBox "Hello " & LastName & ", you are not part of the Management Team."
End Select