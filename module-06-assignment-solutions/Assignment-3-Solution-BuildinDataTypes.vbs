' Assignment Solution

option explicit

Const  SITE_TITLE = "www.GlobalETraining.com" 

'**********************************************************************

Dim EnglishMarks, ScienceMarks, MathMarks, TotalMarks, PercentageMarks

EnglishMarks = 90
ScienceMarks = 95 
MathMarks = 96

TotalMarks = EnglishMarks + ScienceMarks + MathMarks
PercentageMarks = FormatPercent(TotalMarks / 300, 3)


MsgBox "Total number of Marks : " & TotalMarks & vbNewLine & vbNewLine & "Percentage of Marks : " & PercentageMarks, 64, "Marks :"
