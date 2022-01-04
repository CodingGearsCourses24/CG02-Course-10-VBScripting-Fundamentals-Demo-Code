' **********************************************************************
' http://www.CodingGears.com
' FSO
' Assignment 3 - Solution
' **********************************************************************
' Check if the drive has the requried free space

option explicit
Const  SITE_TITLE = ">> CodingGears.io"

Dim drive
Dim oFSO, oDrive
Dim oDrive_total_space, oDrive_free_space
Dim minimum_free_space_in_gb

drive = InputBox("Enter the first drive letter: ", SITE_TITLE, "Drive?" ) 
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oDrive = oFSO.GetDrive(drive)

minimum_free_space_in_gb = InputBox("Minimum Free Space Requirement in GB: ", SITE_TITLE, "Free Space Threshold:" ) 

If oFSO.DriveExists(drive)="True" Then
	oDrive_free_space = ((oDrive.FreeSpace/1024)/1024)/1024
	oDrive_total_space = ((oDrive.TotalSize/1024)/1024)/1024
	
	MsgBox "Free Space (GB): " & FormatNumber(oDrive_free_space, 2),0,drive
	MsgBox "Required Free Space (GB) : " & minimum_free_space_in_gb,0,drive

	If oDrive_free_space > CInt(minimum_free_space_in_gb) Then
		MsgBox "The drive " & drive & " met the minimum free space requirement.",0,"<<< Result >>>"
	Else
		MsgBox "The drive: " & drive & " does not meet the minimum free space requirement.", 0, "<<< ERROR >>>"
	End If
Else
	MsgBox "We did not find the drive: " & drive, 0, "<<< ERROR >>>"
	WScript.Quit
End If