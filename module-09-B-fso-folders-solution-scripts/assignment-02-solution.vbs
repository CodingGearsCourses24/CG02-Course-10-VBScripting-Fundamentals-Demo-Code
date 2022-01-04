' **********************************************************************
' http://www.CodingGears.com
' FSO Folders
' Recursively display subfolders
' Assignment 2 - Solution
' **********************************************************************

Const  SITE_TITLE = " > > >  CodingGears.io < < <"
Dim oFolder

folder1 = InputBox("Enter the full folder path: ", SITE_TITLE, "Full Path?" ) 

set oFSO =createobject("Scripting.Filesystemobject")

If oFSO.FolderExists(folder1) Then
    set oFolder = oFSO.GetFolder(folder1)
    ListSubFolders oFolder
else
	MsgBox folder1 & " - Path Not Found!!! ", (vbOKOnly + vbExclamation)
end if

' Function
Function ListSubFolders(folder_path)
	For Each subfolder in folder_path.SubFolders
        Wscript.Echo subfolder.Path
        ListSubFolders subfolder
    Next
End Function
