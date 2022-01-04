'------------------------------------------------------------
' We will Explore:
'	object.Drives Property
'		Object is always FileSystemObject 
'		Returns a collection with drive objects
'	object.DriveLetter
'		Object is always a Drive object 
'		Returns the drive letter
'	object.Count Property
'		Object is always a collection 
'		Returns total items in a collection)
'------------------------------------------------------------

option explicit

Dim oFSO, oDrive, cDrives
Dim ListOfDrives

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set cDrives = oFSO.Drives

For Each oDrive in cDrives

	MsgBox "Drive letter : " & oDrive.DriveLetter, 0, "Drive On Your Computer:               "
	
	ListOfDrives = ListOfDrives & "   " & oDrive.DriveLetter
	
Next

MsgBox "Number of Drives : " & cDrives.Count, 0, "Drives On Your Computer:"

MsgBox "Drive letters : " & ListOfDrives, 0, "Drives On Your Computer:"