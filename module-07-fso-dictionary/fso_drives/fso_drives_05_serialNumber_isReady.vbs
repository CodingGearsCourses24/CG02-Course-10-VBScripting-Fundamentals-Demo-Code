'------------------------------------------------------------
' We will Explore:
'				[Drive].SerialNumber
'				[Drive].IsReady
'------------------------------------------------------------

option explicit

Dim oFSO, oDrive, cDrives

Dim disktype

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oDrive = oFSO.GetDrive("C:")

MsgBox "Serial Number : " &  oDrive.SerialNumber, 0, "Serial Number (C: ) : "

MsgBox "Is the drive Ready : " &  oDrive.IsReady, 0, "Drive Information : "