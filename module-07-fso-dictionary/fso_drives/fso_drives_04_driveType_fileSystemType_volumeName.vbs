'------------------------------------------------------------
' We will Explore:
'				[Drive].DriveType
'				[Drive].FileSystem
'				[Drive].VolumeName
'------------------------------------------------------------

option explicit

Dim oFSO, oDrive, cDrives

Dim disktype

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oDrive = oFSO.GetDrive("C:")

Select Case oDrive.DriveType
    Case 0: disktype = "Unknown"
    Case 1: disktype = "Removable"
    Case 2: disktype = "Fixed"
    Case 3: disktype = "Network"
    Case 4: disktype = "CD-ROM"
    Case 5: disktype = "RAM Disk"
End Select

MsgBox "Drive Type : " &  oDrive.DriveType, 0, "Drive Type Number Returned by the drive object: "

MsgBox "Drive Type : " &  disktype, 0, "Drive Type : (Use Friendly Strings) "

MsgBox "FileSystem Type : " &  oDrive.FileSystem, 0, "Drive Information : "

MsgBox "Volume Name : " &  oDrive.VolumeName, 0, "Drive Information : "