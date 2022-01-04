' **********************************************************************
' http://www.CodingGears.com
' FSO Folders
' Assignment 1 - Solution
' **********************************************************************

Const  SITE_TITLE = " > > >  CodingGears.io < < <"

'folder1="D:\Wallpapers\"
folder1 = InputBox("Enter the full folder path: ", SITE_TITLE, "Full Path?" ) 

set oFSO =createobject("Scripting.Filesystemobject")

If oFSO.FolderExists(folder1) Then
    set oFolder = oFSO.GetFolder(folder1)
    For Each subfolder in oFolder.SubFolders
        WScript.Echo subfolder.Path
    Next
else
	MsgBox folder1 & " - Path Not Found!!! ", (vbOKOnly + vbExclamation)
end if