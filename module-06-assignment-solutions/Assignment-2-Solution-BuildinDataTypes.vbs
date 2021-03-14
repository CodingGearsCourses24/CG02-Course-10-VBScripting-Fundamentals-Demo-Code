' Assignment Solution

option explicit

Const  SITE_TITLE = "www.GlobalETraining.com" 

'**********************************************************************



Dim TakeOffDate, ReturnDate

TakeOffDate = Date
ReturnDate = DateAdd("d", 10, TakeOffDate)

MsgBox "Flight take off date : " & TakeOffDate & vbNewLine & "Return Date : " & ReturnDate, 64, "Adam's Flight Details :"
