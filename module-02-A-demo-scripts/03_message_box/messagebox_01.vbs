'=========================================================================
'  Syntax: MsgBox prompt [, buttons ] [, title ]
'
'  vbOKOnly				0	Display OK button only.
'  vbOKCancel			1	Display OK and Cancel buttons.
'  vbAbortRetryIgnore	2	Display Abort, Retry, and Ignore buttons.
'  vbYesNoCancel		3	Display Yes, No, and Cancel buttons.
'  vbYesNo				4	Display Yes and No buttons.
'  vbRetryCancel		5	Display Retry and Cancel buttons.
'  *vbCritical			16	Display Critical Message icon.
'  *vbQuestion			32	Display Warning Query icon.
'  *vbExclamation		48	Display Warning Message icon.
'  *vbInformation		64	Display Information Message icon.
'=========================================================================

MsgBox "I am learning VB", 1, "1 :"
 
' 1 + 48 = 49
MsgBox "I am learning VB", 49, "49 :"


' 3 + 64 = 67
MsgBox "I am learning VB", 67, "67 :"

MsgBox "This course is delivered by CodingGears.io. Do you like the course?", 68, "Your Opinion :"