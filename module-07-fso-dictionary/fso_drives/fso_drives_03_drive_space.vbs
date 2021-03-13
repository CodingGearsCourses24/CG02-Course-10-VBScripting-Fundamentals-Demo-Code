'------------------------------------------------------------
' We will Explore:
'				[FSO].GetDrive(drive_letter)
'				[Drive].AvailableSpace
'				[Drive].FreeSpace
'				[Drive].TotalSize
'------------------------------------------------------------

option explicit

Dim oFSO, oDrive, cDrives

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oDrive = oFSO.GetDrive("C:")

MsgBox "Available space : " &  FormatNumber(((oDrive.AvailableSpace/1024)/1024)/1024, 0) & " GB",0, "C Drive"

MsgBox "Free space : " &  FormatNumber(((oDrive.FreeSpace/1024)/1024)/1024, 0) & " GB",0, "C Drive"

MsgBox "Total Size : " &  FormatNumber(((oDrive.TotalSize/1024)/1024)/1024, 0) & " GB",0, "C Drive"