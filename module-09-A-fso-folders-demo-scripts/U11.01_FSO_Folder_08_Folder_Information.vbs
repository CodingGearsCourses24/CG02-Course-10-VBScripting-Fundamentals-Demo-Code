'------------------------------------------------------------
' We will Explore:
' 					Path 
' 					Name
' 					DateCreated
' 					DateLastAccessed
' 					DateLastModified
' 					Size
'------------------------------------------------------------

Dim fso
Dim ReportData1, ReportData2

Dim folder1, ofolder

folder1="D:\VBScripts_Folders\tmp"

set fso =createobject("Scripting.Filesystemobject")

set ofolder = fso.GetFolder(folder1)

ReportData1 = ofolder.Name

MsgBox ReportData1 & " folder " & " on Drive " & ofolder.Drive, 0, "Report: "

ReportData2 = "DateCreated : " & ofolder.DateCreated
ReportData2 = ReportData2 & "DateLastAccessed : " & ofolder.DateLastAccessed
ReportData2 = ReportData2 & "DateLastModified : " & ofolder.DateLastModified  

MsgBox ReportData2,0, "Folder Info"