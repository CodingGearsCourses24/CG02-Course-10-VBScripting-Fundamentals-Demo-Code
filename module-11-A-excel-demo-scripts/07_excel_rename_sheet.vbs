'------------------------------------------------------------
' http://www.GlobalETraining.com
'------------------------------------------------------------
' We will Explore:
'	- Rename a sheet
'------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet
Dim ObjSheet

Dim sheetName
Dim strDirectory
Dim excelDocName

strDirectory = "D:\GitBucket\vbscripting-course\02-indev\excel\tmp_data"
excelDocName = "Sample5.xlsx"
sheetName = "Data1"

'Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = False

'Workbook Object
Set objWorkbook = objExcel.Workbooks.Open(strDirectory & "\" &  excelDocName) 

objWorkbook.Sheets(sheetName).Name = "Data1-new"

'Save, Close & Quit
objWorkbook.Save
objWorkbook.Close 
objExcel.Quit

msgbox "Done!              ", 0, "Status..."

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing