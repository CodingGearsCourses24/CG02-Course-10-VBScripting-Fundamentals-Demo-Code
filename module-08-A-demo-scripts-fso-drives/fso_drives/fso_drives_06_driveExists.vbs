'------------------------------------------------------------
' http://www.GlobalETraining.com
'------------------------------------------------------------
' We will Explore:
'				[Drive].DriveExists
'------------------------------------------------------------

option explicit

Dim oFSO, drive

drive="G:\"

Set oFSO = CreateObject("Scripting.FileSystemObject")

If oFSO.DriveExists(drive)="True" Then
	MsgBox "We found the drive " & drive,0, "Result"
else
	MsgBox "We did not find the drive " & drive, 0, "Result"
End If