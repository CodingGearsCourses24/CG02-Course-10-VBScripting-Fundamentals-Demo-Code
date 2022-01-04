' **********************************************************************
' http://www.CodingGears.com
' FSO
' Assignment 1 - Solution
' **********************************************************************

option explicit

Dim oFSO, oDrive, cDrives
Dim ListOfDrives
Dim disktype

' Total Drives
Set oFSO = CreateObject("Scripting.FileSystemObject")

Set cDrives = oFSO.Drives

MsgBox "Number of Drives : " & cDrives.Count, 0, "Drives On Your Computer:"


' C: Drive
MsgBox "C: Drive details will follow...", 0, "C:\ Drive"

Set oDrive = oFSO.GetDrive("C:")

Select Case oDrive.DriveType
    Case 0: disktype = "Unknown"
    Case 1: disktype = "Removable"
    Case 2: disktype = "Fixed"
    Case 3: disktype = "Network"
    Case 4: disktype = "CD-ROM"
    Case 5: disktype = "RAM Disk"
End Select

MsgBox "Drive Type : " &  oDrive.DriveType, 0, "C:\ Information: "

MsgBox "Drive Type : " &  disktype, 0, "C:\ Information: "

MsgBox "FileSystem Type : " &  oDrive.FileSystem, 0, "C:\ Information: "

MsgBox "Volume Name : " &  oDrive.VolumeName, 0, "C:\ Information: "