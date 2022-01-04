' ==========================================================================
''  		Syntax: MsgBox prompt [, buttons ] [, title ]
' --------------------------------------------------------------------------
'  Buttons
'  -------
'  vbOKOnly				0	Display OK button only.
'  vbOKCancel			1	Display OK and Cancel buttons.
'  vbAbortRetryIgnore	2	Display Abort, Retry, and Ignore buttons.
'  vbYesNoCancel		3	Display Yes, No, and Cancel buttons.
'  vbYesNo				4	Display Yes and No buttons.
'  vbRetryCancel		5	Display Retry and Cancel buttons.
'
'  vbCritical			16	Display Critical Message icon.
'  vbQuestion			32	Display Warning Query icon.
'  vbExclamation		48	Display Warning Message icon.
'  vbInformation		64	Display Information Message icon.
' --------------------------------------------------------------------------
'  Return Values
'  --------------
'  vbOK		1	OK
'  vbCancel	2	Cancel
'  vbAbort	3	Abort
'  vbRetry	4	Retry
'  vbIgnore	5	Ignore
'  vbYes	6	Yes
'  vbNo		7	No
' ==========================================================================

' 3 + 64 = 67 (vbInformation, vbYesNoCancel)
Result = MsgBox("I am learning VB", 67, "Status :")

If Result = 6 Then
	MsgBox "Clicked Yes"
ElseIf Result = vbNo Then
	MsgBox "Clicked No"
ElseIf Result = vbCancel Then
	MsgBox "Clicked Cancel"
End If

