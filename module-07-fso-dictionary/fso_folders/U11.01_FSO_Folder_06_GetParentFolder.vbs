'------------------------------------------------------------
' We will Explore:
'				ParentFolder
'------------------------------------------------------------

Dim oFSO, oFolder, strFolder

strFolder="D:\VBScripts_Folders\tmp\MyFolderData"

Set oFSO=createobject("Scripting.Filesystemobject")

Set oFolder=oFSO.GetFolder(strFolder)

MsgBox "Parent Folder :" &  oFolder.ParentFolder, 0, "Result:"