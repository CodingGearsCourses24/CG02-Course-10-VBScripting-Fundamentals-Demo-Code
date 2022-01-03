'
' Adhoc Script
'
Dim oFSO, sourceFolder, destFolder

Set oFSO=createobject("Scripting.Filesystemobject")

sourceFolder="tmp_data"
destFolder="tmp1"


If oFSO.FolderExists(sourceFolder) Then
	If oFSO.FolderExists(destFolder) Then
		oFSO.DeleteFolder( destFolder )
	Else
		MsgBox destFolder & " " & "not found. Nothing to delete.", 0, "Alert :"
	End if
	
	oFSO.CopyFolder sourceFolder, destFolder, True
	
	MsgBox destFolder & " " & "created using the folder tmp_data", 0, "Alert :"
	
Else
	MsgBox sourceFolder & " " & "not found.", 0, "Alert :"
End if

Set oFSO = Nothing