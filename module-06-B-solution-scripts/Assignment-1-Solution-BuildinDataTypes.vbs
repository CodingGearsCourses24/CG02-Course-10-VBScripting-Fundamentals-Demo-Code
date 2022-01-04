' Assignment Solution

option explicit

Const  SITE_TITLE = "www.GlobalETraining.com" 

'**********************************************************************



Dim MyText, arrWords, TotalNoOfWords, text1, text2, length, MyTextNoSpaces

MyText="The quick brown fox jumps over the lazy dog"

'Determine the total number of words.
arrWords = Split(MyText," ")

TotalNoOfWords=UBound(arrWords) + 1

MsgBox "Total number of words : " & TotalNoOfWords, 64, "Word Count :"

'Extract the word jumps & display a message.
text1 = mid(MyText, 21, 5)

MsgBox "Extracted Characters : " & text1, 64, "Extracted :"


'Display a message with the reverse of the word "quick". (Extract the word from the variable MyText, and then reverse)
text2 = mid(MyText, 5, 5)

MsgBox "Extracted Characters : " & StrReverse(text2), 64, "Extracted :"

'Display a message by removing all white spaces from the variable MyText.

MyTextNoSpaces = Join(arrWords,"")

MsgBox "MyText with no spaces : " & MyTextNoSpaces, 64, "MyText (spaces?) :"


'Display the length of MyText variable.
length=len(MyText)

MsgBox "Length of the MYTEXT string : " & length, 64, "Length :"
