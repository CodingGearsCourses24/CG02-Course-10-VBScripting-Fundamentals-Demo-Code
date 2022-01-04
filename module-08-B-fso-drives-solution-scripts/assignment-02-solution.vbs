' **********************************************************************
' http://www.CodingGears.com
' FSO
' Assignment 2 - Solution
' **********************************************************************
' Prompt the user for two drives, 
' and determine the drive with the more free space.

option explicit
Const  SITE_TITLE = ">> CodingGears.io"

Dim drive1, drive2
Dim oFSO, oDrive1, oDrive2
Dim oDrive1_free_space, oDrive2_free_space

Set oFSO = CreateObject("Scripting.FileSystemObject")
drive1 = InputBox("Enter the first drive letter: ", SITE_TITLE, "First Drive?" ) 
drive2 = InputBox("Enter the second drive letter: ", SITE_TITLE, "Second Drive?" ) 
Set oDrive1 = oFSO.GetDrive(drive1)
Set oDrive2 = oFSO.GetDrive(drive2)

If oFSO.DriveExists(drive1)="True" Then
	If oFSO.DriveExists(drive2)="True"  Then
		MsgBox "We found the drives : " & drive1 & " and " & drive2,0, "Check!"
	Else
		MsgBox "We did not find the drive: " & drive2, 0, "<<< ERROR >>>"
		WScript.Quit
	End If
Else
	MsgBox "We did not find the drive: " & drive1, 0, "<<< ERROR >>>"
	WScript.Quit
End If




'Converting free space into GB
oDrive1_free_space = ((oDrive1.FreeSpace/1024)/1024)/1024
oDrive2_free_space = ((oDrive2.FreeSpace/1024)/1024)/1024

If oDrive1_free_space = oDrive2_free_space Then
	MsgBox "Both drives have same amount of free space!"
End If

If oDrive1_free_space > oDrive2_free_space Then
	MsgBox oDrive1 & " has more free space!",0, oDrive1
    MsgBox "Free space (on " & oDrive1 & " drive " & ") : " &  FormatNumber(oDrive1_free_space, 2) & " GB",0, oDrive1
Else
   	MsgBox oDrive2 & " has more free space!",0, oDrive2
    MsgBox "Free space (on " & oDrive2 & " drive " & ") : " &  FormatNumber(oDrive2_free_space, 2) & " GB",0, oDrive2
End If
